import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Download, Play, FileDown, Settings, HelpCircle } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface Application {
  courseName: string;
  memberId: string;
  name: string;
  phone: string;
  isPriority: boolean;
}

interface Member {
  memberId: string;
  name: string;
  phone: string;
  applications: string[];
  winCount: number;
  wonCourses: string[];
  wonGroups: Set<string>;
}

interface CourseResult {
  winners: Application[];
  waiting: Application[];
}

function shuffle<T>(array: T[]): T[] {
  const newArray = [...array];
  let currentIndex = newArray.length, randomIndex;
  while (currentIndex !== 0) {
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex--;
    [newArray[currentIndex], newArray[randomIndex]] = [newArray[randomIndex], newArray[currentIndex]];
  }
  return newArray;
}

export default function App() {
  const [parsedData, setParsedData] = useState<{ courses: Record<string, Application[]>; applicants: Record<string, Member> } | null>(null);
  const [capacities, setCapacities] = useState<Record<string, number>>({});
  const [courseGroups, setCourseGroups] = useState<Record<string, string>>({});
  const [settings, setSettings] = useState({
    maxWins: 3,
    preventZeroWins: true,
    preventDuplicateGroups: true
  });
  const [results, setResults] = useState<Record<string, CourseResult> | null>(null);
  const [analysis, setAnalysis] = useState<{ unselected: Member[]; duplicates: Member[] } | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [activeTab, setActiveTab] = useState<'course' | 'analysis'>('course');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const ab = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(ab, { type: 'array' });
        
        // 1. Try to find and parse '강좌목록' or '정원' sheet for capacities
        const initialCapacities: Record<string, number> = {};
        const initialGroups: Record<string, string> = {};
        const wsCoursesName = wb.SheetNames.find(name => name.includes('강좌목록') || name.includes('정원'));
        if (wsCoursesName) {
          const coursesData = XLSX.utils.sheet_to_json(wb.Sheets[wsCoursesName]);
          coursesData.forEach((row: any) => {
            const courseName = String(row['강좌명'] || row['과목명'] || row[0] || '').trim();
            const cap = parseInt(row['정원'] || row[1], 10);
            const group = String(row['그룹'] || '').trim();
            
            if (courseName && !isNaN(cap)) {
              initialCapacities[courseName] = cap;
              initialGroups[courseName] = group; // Allow empty group
            }
          });
        }

        // 2. Try to find '신청내역' sheet, otherwise use the first sheet (or second if first is courses)
        let wsName = wb.SheetNames.find(name => name.includes('신청내역'));
        if (!wsName) {
          wsName = (wsCoursesName && wb.SheetNames[0] === wsCoursesName && wb.SheetNames.length > 1) 
            ? wb.SheetNames[1] 
            : wb.SheetNames[0];
        }
        const wsApps = wb.Sheets[wsName];
        const appsData = XLSX.utils.sheet_to_json(wsApps);
        
        const allCourses: Record<string, Application[]> = {};
        const allApplicants: Record<string, Member> = {};
        
        appsData.forEach((row: any) => {
          const courseName = String(row['신청과목'] || row['신청강좌'] || row['강좌명'] || row['과목명'] || row[1] || '').trim();
          const memberId = String(row['회원번호'] || row[2] || '').trim();
          let name = String(row['이름'] || row[3] || '').trim();
          const phone = String(row['전화번호'] || row['연락처'] || row[4] || '').trim();
          
          let isPriority = String(row['우선선정'] || '').trim().toUpperCase() === 'O';
          if (name.includes('*')) {
            isPriority = true;
            name = name.replace(/\*/g, '').trim();
          }
          
          if (!courseName || !memberId) return;
          
          const app: Application = { courseName, memberId, name, phone, isPriority };
          
          if (!allCourses[courseName]) allCourses[courseName] = [];
          allCourses[courseName].push(app);
          
          if (!allApplicants[memberId]) {
            allApplicants[memberId] = {
              memberId,
              name,
              phone,
              applications: [],
              winCount: 0,
              wonCourses: [],
              wonGroups: new Set()
            };
          }
          allApplicants[memberId].applications.push(courseName);
        });
        
        Object.keys(allCourses).forEach(c => {
          if (initialCapacities[c] === undefined) {
            initialCapacities[c] = 15; // Default capacity if not specified in sheet
          }
          if (initialGroups[c] === undefined) {
            initialGroups[c] = ''; // Default group is empty
          }
        });
        
        setParsedData({ courses: allCourses, applicants: allApplicants });
        setCapacities(initialCapacities);
        setCourseGroups(initialGroups);
        setResults(null);
        setAnalysis(null);
      } catch (err: any) {
        alert(err.message || '엑셀 파일 파싱 중 오류가 발생했습니다.');
      }
    };
    reader.readAsArrayBuffer(file);
    
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleCapacityChange = (course: string, value: string) => {
    const num = parseInt(value, 10);
    if (!isNaN(num) && num > 0) {
      setCapacities(prev => ({ ...prev, [course]: num }));
    }
  };

  const handleGroupChange = (course: string, value: string) => {
    setCourseGroups(prev => ({ ...prev, [course]: value }));
  };

  const runLottery = () => {
    if (!parsedData) return;
    
    setIsProcessing(true);
    setResults(null);
    setAnalysis(null);
    
    setTimeout(() => {
      const members: Record<string, Member> = JSON.parse(JSON.stringify(parsedData.applicants));
      // Re-initialize Sets because JSON.parse loses them
      Object.keys(members).forEach(id => {
        members[id].wonGroups = new Set();
      });
      
      const newResults: Record<string, CourseResult> = {};
      Object.keys(parsedData.courses).forEach(c => {
        newResults[c] = { winners: [], waiting: [] };
      });
      
      const maxWinsLimit = settings.maxWins === 0 ? Infinity : settings.maxWins;
      
      // 1. Priority Selection
      Object.keys(parsedData.courses).forEach(course => {
        const priorities = parsedData.courses[course].filter(a => a.isPriority);
        priorities.forEach(p => {
          const group = courseGroups[course] || course;
          const member = members[p.memberId];
          if (member.winCount < maxWinsLimit && (!settings.preventDuplicateGroups || !member.wonGroups.has(group))) {
            newResults[course].winners.push(p);
            member.winCount++;
            member.wonCourses.push(course);
            member.wonGroups.add(group);
          }
        });
      });
      
      // 2. Smart Sequential Drawing
      let memberIds = Object.keys(members);
      memberIds = shuffle(memberIds);
      
      const maxApplications = Math.max(...Object.values(members).map(m => m.applications.length));
      const maxRounds = settings.maxWins === 0 ? maxApplications : settings.maxWins;
      
      for (let round = 1; round <= maxRounds; round++) {
        memberIds.forEach(mId => {
          const member = members[mId];
          if (member.winCount >= maxWinsLimit) return;
          
          const myApps = shuffle([...member.applications]);
          for (const course of myApps) {
            if (member.winCount >= maxWinsLimit) break;
            if (newResults[course].winners.length >= (capacities[course] || 15)) continue;
            
            const group = courseGroups[course] || course;
            if (settings.preventDuplicateGroups && member.wonGroups.has(group)) continue;
            
            if ((round === 1 && member.winCount === 0) || round > 1) {
              const app = parsedData.courses[course].find(a => a.memberId === mId);
              if (app && !newResults[course].winners.some(w => w.memberId === mId)) {
                newResults[course].winners.push(app);
                member.winCount++;
                member.wonCourses.push(course);
                member.wonGroups.add(group);
                break;
              }
            }
          }
        });
      }
      
      // 3. Waiting List
      Object.keys(parsedData.courses).forEach(course => {
        const winnersSet = new Set(newResults[course].winners.map(w => w.memberId));
        const allApplicants = parsedData.courses[course];
        
        // Separate those who won in the same group from those who didn't
        const waiting: Application[] = [];
        const movedToBack: Application[] = [];
        
        allApplicants.forEach(a => {
            if (winnersSet.has(a.memberId)) return;
            
            const member = members[a.memberId];
            const group = courseGroups[course] || course;
            
            if (settings.preventDuplicateGroups && member.wonGroups.has(group)) {
                movedToBack.push(a);
            } else {
                waiting.push(a);
            }
        });
        
        newResults[course].waiting = [...shuffle(waiting), ...shuffle(movedToBack)];
      });
      
      // 4. Guaranteed Selection (Force assign 0-win members)
      if (settings.preventZeroWins) {
        Object.values(members).forEach(member => {
          if (member.winCount === 0 && member.applications.length > 0) {
              // Pick the first application for now
              const course = member.applications[0];
              const app = parsedData.courses[course].find(a => a.memberId === member.memberId);
              if (app) {
                  newResults[course].winners.push(app);
                  member.winCount++;
                  member.wonCourses.push(course);
                  member.wonGroups.add(courseGroups[course] || course);
                  
                  // Remove from waiting list if present
                  newResults[course].waiting = newResults[course].waiting.filter(w => w.memberId !== member.memberId);
              }
          }
        });
      }
      
      // 5. Analysis
      const unselected: Member[] = [];
      const duplicates: Member[] = [];
      Object.values(members).forEach(m => {
        if (m.winCount === 0) unselected.push(m);
        if (m.winCount >= 2) duplicates.push(m);
      });
      
      setResults(newResults);
      setAnalysis({ unselected, duplicates });
      setIsProcessing(false);
    }, 1200);
  };

  const downloadFullExcel = () => {
    if (!results || !analysis) return;
    
    const wb = XLSX.utils.book_new();
    const sheetNames: string[] = [];
    
    // 1. 목차 시트 (Placeholder for now, will fill later)
    const tocWs = XLSX.utils.aoa_to_sheet([["목차"]]);
    XLSX.utils.book_append_sheet(wb, tocWs, "목차");
    
    // 2. 통합 결과 시트
    const allData = [["강좌명", "상태", "순위", "이름", "회원번호", "전화번호", "비고"]];
    for (const course in results) {
      const res = results[course];
      res.winners.forEach((w, i) => allData.push([course, "선정", String(i + 1), w.name, w.memberId, w.phone, w.isPriority ? "우선" : ""]));
      res.waiting.forEach((w, i) => allData.push([course, "대기", String(i + 1), w.name, w.memberId, w.phone, ""]));
    }
    const allWs = XLSX.utils.aoa_to_sheet(allData);
    XLSX.utils.book_append_sheet(wb, allWs, "통합추첨결과");
    sheetNames.push("통합추첨결과");
    
    // 3. 강좌별 시트
    for (const course in results) {
      const res = results[course];
      const data = [["순위", "상태", "이름", "회원번호", "전화번호", "비고"]];
      res.winners.forEach((w, i) => data.push([String(i + 1), "선정", w.name, w.memberId, w.phone, w.isPriority ? "우선" : ""]));
      res.waiting.forEach((w, i) => data.push([String(i + 1), "대기", w.name, w.memberId, w.phone, ""]));
      const ws = XLSX.utils.aoa_to_sheet(data);
      const sheetName = course.substring(0, 30).replace(/[\\/?*[\]]/g, "");
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      sheetNames.push(sheetName);
      
      // Auto-fit column widths
      const cols = data[0].map(() => ({ wch: 15 }));
      ws['!cols'] = cols;
    }
    
    // 4. 미선정자 시트
    if (analysis.unselected.length > 0) {
      const unData = [["순번", "이름", "회원번호", "전화번호", "상태"]];
      analysis.unselected.forEach((m, i) => unData.push([String(i + 1), m.name, m.memberId, m.phone, "전체미선정"]));
      const unWs = XLSX.utils.aoa_to_sheet(unData);
      XLSX.utils.book_append_sheet(wb, unWs, "미선정자명단");
      sheetNames.push("미선정자명단");
    }
    
    // 5. 목차 시트 채우기
    const tocData = [["목차"]];
    sheetNames.forEach(name => tocData.push([name]));
    XLSX.utils.sheet_add_aoa(tocWs, tocData);
    for (let i = 1; i <= sheetNames.length; i++) {
        const cellAddress = XLSX.utils.encode_cell({r: i, c: 0});
        const sheetName = sheetNames[i-1];
        // Escape single quotes in sheet name for Excel formula
        const escapedSheetName = sheetName.replace(/'/g, "''");
        tocWs[cellAddress] = { f: `HYPERLINK("#'${escapedSheetName}'!A1", "${sheetName}")` };
    }
    
    XLSX.writeFile(wb, `복지관_추첨결과_${new Date().toLocaleDateString().replace(/\./g, '').replace(/\s/g, '_')}.xlsx`);
  };

  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    
    // 강좌목록 시트 추가
    const wsCourses = XLSX.utils.json_to_sheet([
      { '강좌명': '요가 1반', '정원': 20, '그룹': '요가' },
      { '강좌명': '요가 2반', '정원': 20, '그룹': '요가' },
      { '강좌명': '필라테스 A', '정원': 15, '그룹': '필라테스' },
      { '강좌명': '필라테스 B', '정원': 15, '그룹': '필라테스' },
      { '강좌명': '인기강좌', '정원': 5, '그룹': '' },
    ]);
    XLSX.utils.book_append_sheet(wb, wsCourses, "강좌목록");

    const wsApps = XLSX.utils.json_to_sheet([
      { '순번': 1, '신청과목': '요가 1반', '회원번호': '1001', '이름': '홍길동', '전화번호': '010-1234-5678', '수강신청상태': '신청완료', '수강진행률': '0%', '등록일': '2023-10-01', '우선선정': 'O' },
      { '순번': 2, '신청과목': '필라테스 A', '회원번호': '1001', '이름': '홍길동', '전화번호': '010-1234-5678', '수강신청상태': '신청완료', '수강진행률': '0%', '등록일': '2023-10-01', '우선선정': '' },
      { '순번': 3, '신청과목': '요가 1반', '회원번호': '1002', '이름': '김철수', '전화번호': '010-9876-5432', '수강신청상태': '신청완료', '수강진행률': '0%', '등록일': '2023-10-02', '우선선정': '' },
      { '순번': 4, '신청과목': '요가 2반', '회원번호': '1002', '이름': '김철수', '전화번호': '010-9876-5432', '수강신청상태': '신청완료', '수강진행률': '0%', '등록일': '2023-10-02', '우선선정': '' },
      { '순번': 5, '신청과목': '인기강좌', '회원번호': '1003', '이름': '이영희', '전화번호': '010-1111-2222', '수강신청상태': '신청완료', '수강진행률': '0%', '등록일': '2023-10-03', '우선선정': '' },
    ]);
    XLSX.utils.book_append_sheet(wb, wsApps, "신청내역");
    XLSX.writeFile(wb, "추첨양식.xlsx");
  };

  return (
    <div className="min-h-screen bg-slate-100 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-7xl mx-auto">
        <header className="flex flex-col md:flex-row justify-between items-center mb-8 gap-4">
          <div>
            <h1 className="text-3xl font-extrabold text-slate-900 leading-tight flex items-center gap-2">
              🎓 강좌 추첨 시스템 <span className="text-blue-600 text-lg ml-2">v3.5</span>
            </h1>
            <p className="text-slate-500 font-medium italic mt-1">고양시대화노인종합복지관 통합 추첨 관리 도구</p>
          </div>
          <div className="flex flex-wrap gap-2 justify-center">
            <button onClick={downloadTemplate} className="bg-white px-3 py-2 rounded shadow-sm border border-slate-200 text-slate-700 font-bold hover:bg-slate-50 flex items-center gap-1 text-sm">
              <Download className="w-4 h-4" /> 양식 다운로드
            </button>
          </div>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          {/* STEP 1 & 2 */}
          <div className="lg:col-span-5 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 border-t-4 border-t-blue-500">
              <h2 className="text-lg font-bold mb-4 flex items-center">
                <span className="bg-blue-100 text-blue-600 w-7 h-7 rounded-full flex items-center justify-center mr-2 text-sm">1</span>
                엑셀 데이터 업로드
              </h2>
              <div 
                className="p-4 border-2 border-dashed border-slate-200 rounded-xl bg-slate-50 text-center relative hover:bg-slate-100 transition-colors cursor-pointer"
                onClick={() => fileInputRef.current?.click()}
              >
                <input type="file" accept=".csv, .xls, .xlsx" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
                <div className="text-4xl mb-2 text-slate-400">📊</div>
                <p className="text-sm font-bold text-slate-600">파일을 선택하거나 여기로 드래그하세요</p>
              </div>
            </div>

            {parsedData && (
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 border-t-4 border-t-amber-500 animate-in fade-in slide-in-from-bottom-4">
                <h2 className="text-lg font-bold mb-4 flex items-center justify-between">
                  <span className="flex items-center">
                    <span className="bg-amber-100 text-amber-600 w-7 h-7 rounded-full flex items-center justify-center mr-2 text-sm">2</span>
                    강좌별 정원 설정
                  </span>
                  <span className="text-[12px] text-red-500 font-bold">신청 60% 미만 시 폐강 주의</span>
                </h2>
                <div className="max-h-[500px] overflow-y-auto space-y-3 pr-2 scrollbar-hide">
                  {Object.keys(parsedData.courses).sort().map(course => {
                    const applicantCount = parsedData.courses[course].length;
                    const cap = capacities[course] || 15;
                    const isUnder60 = applicantCount < cap * 0.6;
                    
                    return (
                      <div key={course} className="flex flex-col p-4 bg-white border border-slate-200 rounded-lg shadow-sm hover:border-blue-300 transition-all gap-2">
                        <div className="flex justify-between items-start gap-2">
                          <div className="min-w-0">
                            <p className="text-base font-bold text-slate-800 leading-tight">{course}</p>
                            <div className="flex items-center gap-2 mt-1">
                              <span className="text-xs text-slate-500 font-medium">신청: <span className="text-blue-600">{applicantCount}명</span></span>
                              {isUnder60 && (
                                <span className="text-xs text-red-500 font-bold ml-2">⚠️ 폐강 가능성 높음</span>
                              )}
                            </div>
                          </div>
                          <div className="flex flex-col gap-2 shrink-0 bg-slate-50 px-3 py-1.5 rounded-md border border-slate-100">
                            <div className="flex items-center gap-2">
                              <span className="text-xs text-slate-500 font-bold w-8">정원</span>
                              <input 
                                type="number" 
                                value={cap} 
                                min="1" 
                                onChange={(e) => handleCapacityChange(course, e.target.value)}
                                className="w-16 border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none"
                              />
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="text-xs text-slate-500 font-bold w-8">그룹</span>
                              <input 
                                type="text" 
                                value={courseGroups[course] || ''} 
                                onChange={(e) => handleGroupChange(course, e.target.value)}
                                className="w-16 border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none"
                              />
                            </div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
                
                <div className="mt-6 p-4 bg-slate-50 border border-slate-200 rounded-lg shadow-sm">
                  <h3 className="text-sm font-bold text-slate-800 mb-3 flex items-center gap-2">
                    <Settings className="w-4 h-4 text-slate-500" />
                    추첨 세팅
                  </h3>
                  <div className="space-y-3">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1">
                        <label className="text-sm text-slate-600 font-medium">인당 최대 선정 수 (0은 무제한)</label>
                        <div className="relative group flex items-center">
                          <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                          <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10">
                            한 명의 신청자가 최대로 당첨될 수 있는 강좌 수입니다.
                          </span>
                        </div>
                      </div>
                      <input 
                        type="number" 
                        value={settings.maxWins} 
                        onChange={(e) => setSettings(s => ({ ...s, maxWins: parseInt(e.target.value) || 0 }))}
                        className="w-16 border border-slate-300 rounded px-2 py-1 text-sm text-center"
                        min="0"
                      />
                    </div>
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1">
                        <label className="text-sm text-slate-600 font-medium">강좌 미선정 방지</label>
                        <div className="relative group flex items-center">
                          <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                          <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10">
                            신청한 강좌 중 하나도 선정되지 않은 신청자에게 최소 1개의 강좌를 강제로 배정합니다.
                          </span>
                        </div>
                      </div>
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input 
                          type="checkbox" 
                          checked={settings.preventZeroWins} 
                          onChange={(e) => setSettings(s => ({ ...s, preventZeroWins: e.target.checked }))}
                          className="sr-only peer"
                        />
                        <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
                      </label>
                    </div>
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1">
                        <label className="text-sm text-slate-600 font-medium">동일 그룹 중복 선정 방지</label>
                        <div className="relative group flex items-center">
                          <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                          <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10">
                            같은 그룹으로 설정된 강좌들 중에서는 최대 1개만 선정되도록 제한합니다.
                          </span>
                        </div>
                      </div>
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input 
                          type="checkbox" 
                          checked={settings.preventDuplicateGroups} 
                          onChange={(e) => setSettings(s => ({ ...s, preventDuplicateGroups: e.target.checked }))}
                          className="sr-only peer"
                        />
                        <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
                      </label>
                    </div>
                  </div>
                </div>

                <button 
                  onClick={runLottery} 
                  className="w-full bg-slate-900 hover:bg-black text-white font-bold py-4 rounded-xl mt-4 transition-all shadow-lg active:scale-95 text-lg flex items-center justify-center gap-2"
                >
                  <Play className="w-5 h-5 fill-current" />
                  강좌 통합 추첨 시작
                </button>
              </div>
            )}
          </div>

          {/* STEP 3: Results */}
          <div className="lg:col-span-7">
            {!parsedData && !isProcessing && !results && (
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 flex flex-col items-center justify-center text-center opacity-60 min-h-[600px]">
                <div className="text-8xl mb-6">🎰</div>
                <h3 className="text-xl font-bold mb-2">강좌 추첨 시스템</h3>
                <p className="text-slate-500 text-sm max-w-sm">
                  데이터를 업로드하면 자동으로 강좌를 분류하며,<br/>중복 당첨 및 미선정자를 관리합니다.
                </p>
              </div>
            )}

            {isProcessing && (
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 flex flex-col items-center justify-center min-h-[600px]">
                <div className="text-7xl mb-6 animate-[spin_3s_linear_infinite]">🎲</div>
                <h3 className="text-2xl font-bold text-blue-600">데이터 정합성 검증 및 추첨 중...</h3>
              </div>
            )}

            {results && analysis && !isProcessing && (
              <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4">
                <div className="bg-white rounded-xl shadow-xl border-b-2 border-slate-100 p-6 flex flex-col sm:flex-row justify-between items-center sticky top-0 z-10 gap-4">
                  <div>
                    <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
                      🏆 통합 추첨 결과
                    </h2>
                    <p className="text-sm text-slate-500 font-medium mt-1">
                      전체 강좌: {Object.keys(results).length}개 | 총 신청인원: {Object.keys(parsedData!.applicants).length}명
                    </p>
                  </div>
                  <button onClick={downloadFullExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2 px-6 rounded-lg transition-all shadow-md flex items-center shrink-0 gap-2">
                    <FileDown className="w-5 h-5" />
                    통합 엑셀 저장
                  </button>
                </div>
                
                {/* Tabs */}
                <div className="flex gap-2 bg-slate-200 p-1 rounded-lg">
                  <button 
                    onClick={() => setActiveTab('course')} 
                    className={cn("flex-1 py-2 text-xs font-bold rounded transition-colors", activeTab === 'course' ? "bg-white shadow-sm text-slate-900" : "text-slate-600 hover:bg-slate-300/50")}
                  >
                    강좌별 결과
                  </button>
                  <button 
                    onClick={() => setActiveTab('analysis')} 
                    className={cn("flex-1 py-2 text-xs font-bold rounded transition-colors", activeTab === 'analysis' ? "bg-white shadow-sm text-slate-900" : "text-slate-600 hover:bg-slate-300/50")}
                  >
                    중복/미선정 관리
                  </button>
                </div>

                {/* Course View */}
                {activeTab === 'course' && (
                  <div className="space-y-8">
                    {Object.entries(results).map(([course, res]: [string, CourseResult]) => {
                      const totalApplicants = parsedData!.courses[course].length;
                      const capacity = capacities[course] || 15;
                      const ratio = Math.round((totalApplicants / capacity) * 100);
                      const barWidth = Math.min(ratio, 100);
                      
                      let barColor = "bg-blue-500";
                      if (ratio >= 100) barColor = "bg-red-500";
                      else if (ratio < 60) barColor = "bg-amber-500";

                      return (
                        <div key={course} className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 border-l-8 border-l-slate-800 overflow-hidden">
                          <div className="mb-5 flex flex-col gap-3">
                            <div className="flex justify-between items-center">
                              <h3 className="text-xl font-bold text-slate-900 leading-tight">{course}</h3>
                              <div className="flex gap-2">
                                <span className="text-xs bg-emerald-100 text-emerald-700 font-bold px-2 py-1 rounded">선정 {res.winners.length}</span>
                                <span className="text-xs bg-orange-100 text-orange-700 font-bold px-2 py-1 rounded">대기 {res.waiting.length}</span>
                              </div>
                            </div>
                            <div className="flex flex-col gap-1">
                              <div className="flex justify-between items-center text-sm font-medium text-slate-500">
                                <span>신청률: {ratio}% ({totalApplicants}명 / 정원 {capacity}명)</span>
                                {ratio < 60 && <span className="text-red-500 font-bold">⚠️ 폐강 주의</span>}
                              </div>
                              <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
                                <div className={cn("h-full transition-all duration-500", barColor)} style={{ width: `${barWidth}%` }}></div>
                              </div>
                            </div>
                          </div>
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                            <div>
                              <h4 className="text-sm font-bold text-emerald-800 mb-2">선정 명단</h4>
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                {res.winners.map((w, idx) => (
                                  <div key={idx} className={cn("p-3 rounded flex justify-between items-start shadow-sm", w.isPriority ? "bg-blue-50 border-l-4 border-blue-500" : "bg-emerald-50 border-l-4 border-emerald-500")}>
                                    <div className="flex flex-col">
                                      <span className="text-sm md:text-base text-slate-800 font-bold">{idx + 1}. {w.name}</span>
                                      <span className="text-xs text-slate-500 mt-0.5">{w.memberId}</span>
                                    </div>
                                  </div>
                                ))}
                                {res.winners.length === 0 && <p className="text-sm text-slate-400 italic">선정자가 없습니다.</p>}
                              </div>
                            </div>
                            <div>
                              <h4 className="text-sm font-bold text-orange-800 mb-2">대기 순번</h4>
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                {res.waiting.map((w, idx) => (
                                  <div key={idx} className="bg-orange-50 border-l-4 border-orange-500 p-3 rounded shadow-sm flex flex-col justify-center">
                                    <span className="text-sm md:text-base text-slate-700 font-bold">{idx + 1}. {w.name}</span>
                                    <span className="text-xs text-slate-500 mt-0.5">{w.memberId}</span>
                                  </div>
                                ))}
                                {res.waiting.length === 0 && <p className="text-sm text-slate-400 italic">대기자가 없습니다.</p>}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* Analysis View */}
                {activeTab === 'analysis' && (
                  <div className="space-y-8">
                    {/* Duplicates */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-8 border-l-8 border-l-red-500">
                      <div className="mb-6 flex justify-between items-end">
                        <div>
                          <h3 className="text-xl font-bold text-slate-900 flex items-center">
                            <span className="mr-2">👯</span> 중복 선정자 명단 (2개 이상)
                          </h3>
                          <p className="text-sm text-slate-500 mt-1 italic">선정이 편중된 분들의 현황입니다.</p>
                        </div>
                        <span className="text-red-600 font-bold text-sm">{analysis.duplicates.length}명</span>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                        {analysis.duplicates.length > 0 ? analysis.duplicates.map((m, idx) => (
                          <div key={idx} className="bg-red-50 border-l-4 border-red-500 p-3 rounded flex flex-col shadow-sm">
                            <div className="flex justify-between">
                              <span className="text-sm font-bold text-slate-800">{m.name}</span>
                              <span className="text-xs bg-red-100 text-red-600 px-1.5 rounded flex items-center">{m.winCount}개</span>
                            </div>
                            <span className="text-xs text-slate-500 mt-1">{m.wonCourses.join(', ')}</span>
                          </div>
                        )) : (
                          <p className="col-span-full text-center text-slate-400 py-4 italic">중복 선정자가 없습니다.</p>
                        )}
                      </div>
                    </div>

                    {/* Unselected */}
                    <div className="bg-slate-50 rounded-xl shadow-sm border border-slate-200 p-8 border-l-8 border-l-slate-400">
                      <div className="mb-6 flex justify-between items-end">
                        <div>
                          <h3 className="text-xl font-bold text-slate-900 flex items-center">
                            <span className="mr-2">👤</span> 최종 미선정자 명단
                          </h3>
                          <p className="text-sm text-slate-500 mt-1 italic">단 하나의 강좌도 선정되지 못한 분들입니다.</p>
                        </div>
                        <span className="text-slate-600 font-bold text-sm">{analysis.unselected.length}명</span>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                        {analysis.unselected.length > 0 ? analysis.unselected.map((m, idx) => (
                          <div key={idx} className="bg-slate-100 border-l-4 border-slate-500 p-3 rounded flex flex-col justify-center shadow-sm">
                            <span className="text-sm font-bold text-slate-700">{idx + 1}. {m.name}</span>
                            <span className="text-xs text-slate-500">{m.memberId}</span>
                          </div>
                        )) : (
                          <p className="col-span-full text-center text-slate-400 py-4 italic">미선정자가 없습니다.</p>
                        )}
                      </div>
                    </div>
                  </div>
                )}
                
                <div className="pb-20"></div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
