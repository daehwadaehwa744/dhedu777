import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Download, Play, FileDown, Settings, HelpCircle, Loader2, X } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function parseTimeStr(timeStr: string) {
  if (!timeStr) return null;
  try { return JSON.parse(timeStr); } catch { return null; }
}

function timeToMinutes(tStr: string) {
  const parts = tStr.split(':');
  if (parts.length >= 2) {
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    if (!isNaN(h) && !isNaN(m)) return h * 60 + m;
  }
  return 0;
}

function checkTimeConflict(t1Str: string, wonTimes: Set<string>) {
  if (!t1Str) return false;
  const t1 = parseTimeStr(t1Str);
  if (!t1 || !t1.day) return false;

  const days1 = t1.day.split(',').map((d: string) => d.trim()).filter(Boolean);

  for (const wt of wonTimes) {
    const t2 = parseTimeStr(wt);
    if (!t2 || !t2.day) continue;
    
    const days2 = t2.day.split(',').map((d: string) => d.trim()).filter(Boolean);
    const hasCommonDay = days1.some((d: string) => days2.includes(d));
    if (!hasCommonDay) continue;

    if (!t1.start || !t1.end || !t2.start || !t2.end) {
      return true; // assume conflict if day matches but time is missing
    }

    const t1s = timeToMinutes(t1.start);
    const t1e = timeToMinutes(t1.end);
    const t2s = timeToMinutes(t2.start);
    const t2e = timeToMinutes(t2.end);

    if (t1s < t2e && t2s < t1e) {
      return true;
    }
  }
  return false;
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
  wonTimes: Set<string>;
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
  const [courseTimes, setCourseTimes] = useState<Record<string, string>>({});
  const [settings, setSettings] = useState({
    maxWins: 3,
    fairDistribution: true,
    preventZeroWins: true,
    preventDuplicateGroups: true,
    preventDuplicateTimes: true,
    fillUnderfilled: true
  });
  const [results, setResults] = useState<Record<string, CourseResult> | null>(null);
  const [analysis, setAnalysis] = useState<{ unselected: Member[]; duplicates: Member[] } | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [activeTab, setActiveTab] = useState<'course' | 'duplicates' | 'unselected'>('course');
  const [showSettingsModal, setShowSettingsModal] = useState(false);
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
        const initialTimes: Record<string, string> = {};
        const wsCoursesName = wb.SheetNames.find(name => name.includes('강좌목록') || name.includes('정원'));
        if (wsCoursesName) {
          const coursesData = XLSX.utils.sheet_to_json(wb.Sheets[wsCoursesName]);
          coursesData.forEach((row: any) => {
            const courseName = String(row['강좌명'] || row['과목명'] || row[0] || '').trim();
            const cap = parseInt(row['정원'] || row[1], 10);
            const group = String(row['그룹'] || '').trim();
            const dayInfo = String(row['요일'] || '').trim();
            const startTime = String(row['시작 시간'] || row['시작시간'] || '').trim();
            const endTime = String(row['종료 시간'] || row['종료시간'] || '').trim();
            
            if (courseName && !isNaN(cap)) {
              initialCapacities[courseName] = cap;
              initialGroups[courseName] = group; // Allow empty group
              initialTimes[courseName] = JSON.stringify({ day: dayInfo, start: startTime, end: endTime });
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
              wonGroups: new Set(),
              wonTimes: new Set()
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
          if (initialTimes[c] === undefined) {
            initialTimes[c] = JSON.stringify({ day: '', start: '', end: '' }); // Default time is empty
          }
        });
        
        setParsedData({ courses: allCourses, applicants: allApplicants });
        setCapacities(initialCapacities);
        setCourseGroups(initialGroups);
        setCourseTimes(initialTimes);
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

  const handleTimeChange = (course: string, field: 'day'|'start'|'end', val: string) => {
    setCourseTimes(prev => {
      const t = parseTimeStr(prev[course] || '{}') || { day: '', start: '', end: '' };
      t[field] = val;
      return { ...prev, [course]: JSON.stringify(t) };
    });
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
        members[id].wonTimes = new Set();
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
          const time = courseTimes[course] || '';
          const member = members[p.memberId];
          
          const noGroupConflict = !settings.preventDuplicateGroups || !member.wonGroups.has(group);
          const noTimeConflict = !settings.preventDuplicateTimes || time === '' || !checkTimeConflict(time, member.wonTimes);

          if (member.winCount < maxWinsLimit && noGroupConflict && noTimeConflict) {
            newResults[course].winners.push(p);
            member.winCount++;
            member.wonCourses.push(course);
            member.wonGroups.add(group);
            if (time) member.wonTimes.add(time);
          }
        });
      });
      
      // 2. Base Allocation (Guaranteed 1st Win Fairly)
      // "가장 선택지가 적은(여유가 없는) 사람"부터 배정하여 0건 당첨을 최소화하는 최적 배정 알고리즘
      let memberIds = Object.keys(members);
      memberIds = shuffle(memberIds);
      
      const maxApplications = Math.max(...Object.values(members).map(m => m.applications.length));
      const maxRounds = settings.maxWins === 0 ? maxApplications : settings.maxWins;
      
      if (settings.preventZeroWins) {
        let changed = true;
        while (changed) {
          changed = false;
          
          const candidates: { mId: string; availableCourses: string[] }[] = [];
          
          memberIds.forEach(mId => {
            const member = members[mId];
            if (member.winCount === 0) {
              const available = member.applications.filter(course => {
                const group = courseGroups[course] || course;
                const time = courseTimes[course] || '';
                const hasCapacity = newResults[course].winners.length < (capacities[course] || 15);
                const noGroupConflict = !settings.preventDuplicateGroups || !member.wonGroups.has(group);
                const noTimeConflict = !settings.preventDuplicateTimes || time === '' || !checkTimeConflict(time, member.wonTimes);
                // Also check if they are already in winners (shouldn't happen but safe)
                const notAlreadyWinner = !newResults[course].winners.some(w => w.memberId === mId);
                return hasCapacity && noGroupConflict && noTimeConflict && notAlreadyWinner;
              });
              
              if (available.length > 0) {
                candidates.push({ mId, availableCourses: available });
              }
            }
          });
          
          if (candidates.length > 0) {
            // Find the minimum number of available courses
            const minAvailable = Math.min(...candidates.map(c => c.availableCourses.length));
            
            // Get all members who are most constrained
            const mostConstrained = candidates.filter(c => c.availableCourses.length === minAvailable);
            
            // Pick one randomly
            const chosen = mostConstrained[Math.floor(Math.random() * mostConstrained.length)];
            const member = members[chosen.mId];
            
            // Pick a random available course
            const course = chosen.availableCourses[Math.floor(Math.random() * chosen.availableCourses.length)];
            
            const group = courseGroups[course] || course;
            const time = courseTimes[course] || '';
            const app = parsedData.courses[course].find(a => a.memberId === chosen.mId);
            
            if (app) {
              newResults[course].winners.push(app);
              member.winCount++;
              member.wonCourses.push(course);
              member.wonGroups.add(group);
              if (time) member.wonTimes.add(time);
              changed = true;
            }
          }
        }
      }

      // 3. Drawing Strategy
      if (settings.fairDistribution) {
        // Smart Sequential Drawing (공평 분배: 사람을 기준으로 라운드 로빈 순회)
        for (let round = 1; round <= maxRounds; round++) {
          memberIds = shuffle(memberIds);
          memberIds.forEach(mId => {
            const member = members[mId];
            if (member.winCount >= maxWinsLimit) return;
            
            const myApps = shuffle([...member.applications]);
            for (const course of myApps) {
              if (member.winCount >= maxWinsLimit) break;
              if (newResults[course].winners.length >= (capacities[course] || 15)) continue;
              
              const group = courseGroups[course] || course;
              const time = courseTimes[course] || '';
              if (settings.preventDuplicateGroups && member.wonGroups.has(group)) continue;
              if (settings.preventDuplicateTimes && time !== '' && checkTimeConflict(time, member.wonTimes)) continue;
              
              const app = parsedData.courses[course].find(a => a.memberId === mId);
              if (app && !newResults[course].winners.some(w => w.memberId === mId)) {
                newResults[course].winners.push(app);
                member.winCount++;
                member.wonCourses.push(course);
                member.wonGroups.add(group);
                if (time) member.wonTimes.add(time);
                break;
              }
            }
          });
        }
      } else {
        // Independent Course Drawing (강좌별 독립 무작위 추첨)
        const courseList = shuffle(Object.keys(parsedData.courses));
        for (const course of courseList) {
          const applicants = shuffle([...parsedData.courses[course]]);
          for (const app of applicants) {
            const mId = app.memberId;
            const member = members[mId];
            if (newResults[course].winners.length >= (capacities[course] || 15)) break; // 강좌 정원 참
            
            if (member.winCount >= maxWinsLimit) continue;
            
            // 이미 이 강좌에 뽑힌 적 있는지 확인
            if (newResults[course].winners.some(w => w.memberId === mId)) continue;
            
            const group = courseGroups[course] || course;
            const time = courseTimes[course] || '';
            if (settings.preventDuplicateGroups && member.wonGroups.has(group)) continue;
            if (settings.preventDuplicateTimes && time !== '' && checkTimeConflict(time, member.wonTimes)) continue;
            
            newResults[course].winners.push(app);
            member.winCount++;
            member.wonCourses.push(course);
            member.wonGroups.add(group);
            if (time) member.wonTimes.add(time);
          }
        }
      }

      // 3.5. Forced Re-allocation (Robin Hood Logic)
      // 정원 초과로 0건 배정된 수강생을 위해, 다기수 당첨자(2건 이상)의 자리를 회수하여 배정합니다.
      if (settings.preventZeroWins) {
        const zeroWinIdList = Object.keys(members).filter(mId => members[mId].winCount === 0 && members[mId].applications.length > 0);
        const shuffledZeroWins = shuffle(zeroWinIdList);

        for (const mId of shuffledZeroWins) {
          const zeroMember = members[mId];
          if (zeroMember.winCount > 0) continue; 
          
          let bestVictimInfo: { vId: string, course: string, winCount: number } | null = null;
          
          const myApps = shuffle([...zeroMember.applications]);
          for (const course of myApps) {
            const currentWinners = newResults[course].winners;
            
            for (const winnerApp of currentWinners) {
              const winnerMember = members[winnerApp.memberId];
              // 2개 이상 선정된 수강생 타겟팅
              if (winnerMember.winCount > 1) {
                // 가장 많이 당첨된 사람을 우선 희생자로 선정
                if (!bestVictimInfo || winnerMember.winCount > bestVictimInfo.winCount) {
                  bestVictimInfo = {
                    vId: winnerMember.memberId,
                    course: course,
                    winCount: winnerMember.winCount
                  };
                }
              }
            }
          }
          
          // 희생자를 찾았다면 Swap 수행
          if (bestVictimInfo) {
            const course = bestVictimInfo.course;
            const victimMem = members[bestVictimInfo.vId];
            const zeroMemApp = parsedData.courses[course].find(a => a.memberId === mId);
            
            if (zeroMemApp) {
              // 희생자 목록에서 강좌 제외
              newResults[course].winners = newResults[course].winners.filter(w => w.memberId !== bestVictimInfo.vId);
              victimMem.winCount--;
              victimMem.wonCourses = victimMem.wonCourses.filter(c => c !== course);
              
              // 희생자 그룹/시간셋 최신화
              victimMem.wonGroups.clear();
              victimMem.wonTimes.clear();
              victimMem.wonCourses.forEach(c => {
                victimMem.wonGroups.add(courseGroups[c] || c);
                if (courseTimes[c]) victimMem.wonTimes.add(courseTimes[c]);
              });
              
              // 0건 수강생에게 강좌 배정
              newResults[course].winners.push(zeroMemApp);
              zeroMember.winCount++;
              zeroMember.wonCourses.push(course);
              zeroMember.wonGroups.add(courseGroups[course] || course);
              if (courseTimes[course]) zeroMember.wonTimes.add(courseTimes[course]);
            }
          }
        }
      }

      // 4. Fill Underfilled Courses (Ignore maxWins)
      // 정원 미달 강좌의 경우 추가적으로 채우기
      if (settings.fillUnderfilled) {
        for (let round = 1; round <= maxApplications; round++) {
          memberIds = shuffle(memberIds); // 순서를 다시 섞어 공정성 유지
          memberIds.forEach(mId => {
            const member = members[mId];
            const myApps = shuffle([...member.applications]);
            for (const course of myApps) {
              if (newResults[course].winners.length >= (capacities[course] || 15)) continue;
              
              const group = courseGroups[course] || course;
              const time = courseTimes[course] || '';
              if (settings.preventDuplicateGroups && member.wonGroups.has(group)) continue;
              if (settings.preventDuplicateTimes && time !== '' && checkTimeConflict(time, member.wonTimes)) continue;
              
              const app = parsedData.courses[course].find(a => a.memberId === mId);
              if (app && !newResults[course].winners.some(w => w.memberId === mId)) {
                newResults[course].winners.push(app);
                member.winCount++;
                member.wonCourses.push(course);
                member.wonGroups.add(group);
                if (time) member.wonTimes.add(time);
                break; // 한 라운드당 한 명에게 하나씩만 추가
              }
            }
          });
        }
      }
      
      // 5. Waiting List
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
            const time = courseTimes[course] || '';
            const hasGroupConflict = settings.preventDuplicateGroups && member.wonGroups.has(group);
            const hasTimeConflict = settings.preventDuplicateTimes && time !== '' && checkTimeConflict(time, member.wonTimes);
            
            if (hasGroupConflict || hasTimeConflict) {
                movedToBack.push(a);
            } else {
                waiting.push(a);
            }
        });
        
        newResults[course].waiting = [...shuffle(waiting), ...shuffle(movedToBack)];
      });
      
      // 6. Analysis
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
    const memberSummary: Record<string, { memberId: string, name: string, phone: string, won: string[], waiting: string[] }> = {};
    
    for (const course in results) {
      const res = results[course];
      res.winners.forEach(w => {
        if (!memberSummary[w.memberId]) memberSummary[w.memberId] = { memberId: w.memberId, name: w.name, phone: w.phone, won: [], waiting: [] };
        memberSummary[w.memberId].won.push(course);
      });
      res.waiting.forEach(w => {
        if (!memberSummary[w.memberId]) memberSummary[w.memberId] = { memberId: w.memberId, name: w.name, phone: w.phone, won: [], waiting: [] };
        memberSummary[w.memberId].waiting.push(course);
      });
    }

    let maxWon = 0;
    let maxWait = 0;
    const summaryList = Object.values(memberSummary);
    summaryList.forEach(m => {
      if (m.won.length > maxWon) maxWon = m.won.length;
      if (m.waiting.length > maxWait) maxWait = m.waiting.length;
    });

    const headerRow = ["이름", "회원번호", "전화번호"];
    for (let i = 1; i <= maxWon; i++) headerRow.push(`선정 ${i}`);
    for (let i = 1; i <= maxWait; i++) headerRow.push(`대기 ${i}`);

    const allData = [headerRow];
    
    summaryList.sort((a, b) => a.name.localeCompare(b.name)).forEach(m => {
      const row = [m.name, m.memberId, m.phone];
      for (let i = 0; i < maxWon; i++) {
        row.push(m.won[i] || "");
      }
      for (let i = 0; i < maxWait; i++) {
        row.push(m.waiting[i] || "");
      }
      allData.push(row);
    });

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
      { '강좌명': '요가 1반', '정원': 20, '그룹': '요가', '요일': '월', '시작 시간': '10:00', '종료 시간': '11:00' },
      { '강좌명': '요가 2반', '정원': 20, '그룹': '요가', '요일': '화', '시작 시간': '10:00', '종료 시간': '11:00' },
      { '강좌명': '필라테스 A', '정원': 15, '그룹': '필라테스', '요일': '월', '시작 시간': '10:00', '종료 시간': '11:00' },
      { '강좌명': '필라테스 B', '정원': 15, '그룹': '필라테스', '요일': '수', '시작 시간': '14:00', '종료 시간': '15:00' },
      { '강좌명': '인기강좌', '정원': 5, '그룹': '', '요일': '', '시작 시간': '', '종료 시간': '' },
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

  const renderSettingsContent = () => (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <h2 className="text-lg font-bold mb-4 flex items-center">
          <span className="bg-emerald-100 text-emerald-600 w-7 h-7 rounded-full flex items-center justify-center mr-2 text-sm">A</span>
          추첨 규칙 설정
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          <div className="flex flex-col gap-2 bg-slate-50 p-4 rounded-lg border border-slate-100">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1">
                <label className="text-sm text-slate-600 font-bold">1인당 최대 선정 강좌 수</label>
                <div className="relative group flex items-center">
                  <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                  <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 hidden md:block">
                    한 회원이 우선선정과 무작위 추첨을 포함해 최대로 선정될 수 있는 강좌의 수입니다.
                  </span>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <input 
                  type="number" 
                  value={settings.maxWins} 
                  min="1"
                  onChange={(e) => setSettings(s => ({ ...s, maxWins: parseInt(e.target.value) || 1 }))}
                  className="w-16 border-2 border-slate-200 rounded px-2 py-1 text-base font-bold text-blue-600 focus:border-blue-400 outline-none text-right"
                />
                <span className="text-sm font-medium text-slate-500">개</span>
              </div>
            </div>
            <div className="flex items-center justify-between mt-2">
              <div className="flex items-center gap-1">
                <label className="text-sm text-slate-600 font-medium whitespace-nowrap">공평 분배 (라운드 로빈)</label>
                <div className="relative group flex items-center">
                  <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                  <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 hidden md:block">
                    1인당 1장씩 순서대로 배분합니다. 끄면 강좌별 독립 무작위 추첨으로 변경되어 한 명이 여러 강좌에 선정될 확률이 자연스러워집니다.
                  </span>
                </div>
              </div>
              <label className="relative inline-flex items-center cursor-pointer shrink-0">
                <input 
                  type="checkbox" 
                  checked={settings.fairDistribution} 
                  onChange={(e) => setSettings(s => ({ ...s, fairDistribution: e.target.checked }))}
                  className="sr-only peer"
                />
                <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
              </label>
            </div>
          </div>
          
          <div className="flex flex-col gap-3 justify-center bg-slate-50 p-4 rounded-lg border border-slate-100">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1">
                <label className="text-sm text-slate-600 font-medium whitespace-nowrap">미선정 방지(최소 1개 우선보장)</label>
                <div className="relative group flex items-center">
                  <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                  <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 hidden md:block">
                    모든 강좌에서 떨어진 사람을 모아, 대기열에 있는 강좌 중 자리가 남는 곳에 1순위로 배정합니다.
                  </span>
                </div>
              </div>
              <label className="relative inline-flex items-center cursor-pointer shrink-0">
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
                <label className="text-sm text-slate-600 font-medium whitespace-nowrap">동일 그룹 중복 선정 방지</label>
                <div className="relative group flex items-center">
                  <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                  <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 hidden md:block">
                    같은 그룹명(예: "요가")으로 지정된 강좌는 1인당 1개만 선정되도록 제한합니다.
                  </span>
                </div>
              </div>
              <label className="relative inline-flex items-center cursor-pointer shrink-0">
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
          
          <div className="flex flex-col gap-3 justify-center bg-slate-50 p-4 rounded-lg border border-slate-100">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1">
                <label className="text-sm text-slate-600 font-medium whitespace-nowrap">동일 시간대 중복 선정 방지</label>
                <div className="relative group flex items-center">
                  <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                  <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 hidden md:block">
                    동일한 수업시간을 가진 강좌가 1인당 1개만 선정정되도록 제한합니다.
                  </span>
                </div>
              </div>
              <label className="relative inline-flex items-center cursor-pointer shrink-0">
                <input 
                  type="checkbox" 
                  checked={settings.preventDuplicateTimes} 
                  onChange={(e) => setSettings(s => ({ ...s, preventDuplicateTimes: e.target.checked }))}
                  className="sr-only peer"
                />
                <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
              </label>
            </div>
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-1">
                <label className="text-sm text-slate-600 font-medium whitespace-nowrap">정원 미달 강좌 추가 선정</label>
                <div className="relative group flex items-center">
                  <HelpCircle className="w-3 h-3 text-slate-400 cursor-help" />
                  <span className="absolute left-1/2 -translate-x-1/2 bottom-full mb-2 w-64 p-2 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-10 hidden md:block">
                    정원이 미달된 강좌에 한해, 이미 당첨된 회원이라도 중복 당첨될 수 있도록 추가 배정합니다.
                  </span>
                </div>
              </div>
              <label className="relative inline-flex items-center cursor-pointer shrink-0">
                <input 
                  type="checkbox" 
                  checked={settings.fillUnderfilled} 
                  onChange={(e) => setSettings(s => ({ ...s, fillUnderfilled: e.target.checked }))}
                  className="sr-only peer"
                />
                <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
              </label>
            </div>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-lg font-bold flex items-center">
            <span className="bg-amber-100 text-amber-600 w-7 h-7 rounded-full flex items-center justify-center mr-2 text-sm">B</span>
            강좌별 정원 및 그룹 설정
          </h2>
          <span className="text-sm text-red-500 font-bold">신청 60% 미만 시 폐강 주의</span>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {Object.keys(parsedData?.courses || {}).sort().map(course => {
            const applicantCount = parsedData!.courses[course].length;
            const cap = capacities[course] || 15;
            const isUnder60 = applicantCount < cap * 0.6;
            const timeInfo = parseTimeStr(courseTimes[course] || '{}') || { day: '', start: '', end: '' };
            
            return (
              <div key={course} className="flex flex-col p-4 bg-white border border-slate-200 rounded-lg shadow-sm hover:border-blue-300 transition-all gap-3">
                <div className="flex justify-between items-start">
                  <div className="min-w-0 flex flex-col gap-1 w-full">
                    <p className="text-base font-bold text-slate-800 leading-tight w-full truncate" title={course}>{course}</p>
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-slate-500 font-medium">신청: <span className="text-blue-600">{applicantCount}명</span></span>
                      {isUnder60 && (
                        <span className="text-xs bg-red-100 text-red-600 px-1.5 py-0.5 rounded font-bold whitespace-nowrap">⚠️ 폐강주의</span>
                      )}
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div className="flex flex-col bg-slate-50 px-3 py-2 rounded-md border border-slate-100">
                    <span className="text-xs text-slate-500 font-bold mb-1">정원</span>
                    <input 
                      type="number" 
                      value={cap} 
                      min="1" 
                      onChange={(e) => handleCapacityChange(course, e.target.value)}
                      className="w-full border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none"
                    />
                  </div>
                  <div className="flex flex-col bg-slate-50 px-3 py-2 rounded-md border border-slate-100">
                    <span className="text-xs text-slate-500 font-bold mb-1">그룹</span>
                    <input 
                      type="text" 
                      value={courseGroups[course] || ''} 
                      onChange={(e) => handleGroupChange(course, e.target.value)}
                      className="w-full border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none"
                    />
                  </div>
                </div>
                
                <div className="flex flex-col bg-slate-50 px-3 py-2 rounded-md border border-slate-100">
                  <div className="text-xs text-slate-500 font-bold mb-2">시간대</div>
                  <div className="grid grid-cols-3 gap-2">
                    <div className="flex flex-col gap-1">
                      <span className="text-xs text-slate-400 font-medium">요일</span>
                      <input
                        type="text"
                        placeholder="월,수"
                        value={timeInfo.day}
                        onChange={(e) => handleTimeChange(course, 'day', e.target.value)}
                        className="w-full border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none"
                      />
                    </div>
                    <div className="flex flex-col gap-1">
                      <span className="text-xs text-slate-400 font-medium">시작시간</span>
                      <input
                        type="text"
                        placeholder="09:00"
                        maxLength={5}
                        value={timeInfo.start}
                        onChange={(e) => {
                          let val = e.target.value.replace(/[^0-9:]/g, '');
                          if (val.length === 2 && !val.includes(':') && e.target.value.length > timeInfo.start.length) val += ':';
                          handleTimeChange(course, 'start', val);
                        }}
                        className="w-full border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none placeholder:text-blue-200"
                      />
                    </div>
                    <div className="flex flex-col gap-1">
                      <span className="text-xs text-slate-400 font-medium">종료시간</span>
                      <input
                        type="text"
                        placeholder="11:30"
                        maxLength={5}
                        value={timeInfo.end}
                        onChange={(e) => {
                          let val = e.target.value.replace(/[^0-9:]/g, '');
                          if (val.length === 2 && !val.includes(':') && e.target.value.length > timeInfo.end.length) val += ':';
                          handleTimeChange(course, 'end', val);
                        }}
                        className="w-full border-2 border-white rounded px-2 py-1 text-sm font-bold text-blue-600 focus:border-blue-400 outline-none placeholder:text-blue-200"
                      />
                    </div>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-100 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-7xl mx-auto">
        <header className="flex flex-col md:flex-row justify-between items-center mb-8 gap-4">
          <div>
            <h1 className="text-3xl font-extrabold text-slate-900 leading-tight flex items-center gap-2">
              🎓 강좌 추첨 시스템 <span className="text-blue-600 text-lg ml-2">v4.2</span>
            </h1>
            <p className="text-slate-500 font-medium italic mt-1">노년사회화교육 강좌 추첨 관리 도구</p>
          </div>
          <div className="flex flex-wrap gap-2 justify-center">
            <button onClick={downloadTemplate} className="bg-white px-3 py-2 rounded shadow-sm border border-slate-200 text-slate-700 font-bold hover:bg-slate-50 flex items-center gap-1 text-sm">
              <Download className="w-4 h-4" /> 양식 다운로드
            </button>
          </div>
        </header>

        {!parsedData ? (
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 border-t-4 border-t-blue-500 max-w-3xl mx-auto mt-12 animate-in fade-in">
            <h2 className="text-2xl font-bold mb-6 flex items-center justify-center">
              <span className="bg-blue-100 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center mr-3 text-lg">1</span>
              엑셀 데이터 업로드
            </h2>
            <div 
              className="p-12 border-2 border-dashed border-slate-300 rounded-xl bg-slate-50 text-center relative hover:bg-slate-100 transition-colors cursor-pointer"
              onClick={() => fileInputRef.current?.click()}
            >
              <input type="file" accept=".csv, .xls, .xlsx" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
              <div className="text-6xl mb-4 text-slate-400">📊</div>
              <p className="text-xl font-bold text-slate-600 mb-2">파일을 선택하거나 여기로 드래그하세요</p>
              <p className="text-slate-500">지원 형식: .xlsx, .xls, .csv</p>
            </div>
          </div>
        ) : !results ? (
          <div className="animate-in fade-in space-y-6">
            <div className="flex items-center gap-3 mb-6">
              <span className="bg-amber-100 text-amber-600 w-10 h-10 rounded-full flex items-center justify-center text-xl font-bold">2</span>
              <h2 className="text-2xl font-bold text-slate-800">강좌별 정원 및 추첨 설정</h2>
            </div>
            
            {renderSettingsContent()}

            <div className="flex justify-center mt-8">
              <button 
                onClick={runLottery} 
                disabled={isProcessing}
                className="bg-blue-600 outline-none hover:bg-blue-700 text-white font-bold py-4 px-12 rounded-xl shadow-lg border border-transparent transition-all disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center text-xl gap-2 w-full md:w-auto min-w-[300px]"
              >
                {isProcessing ? (
                  <>
                    <Loader2 className="w-6 h-6 animate-spin" />
                    추첨 진행 중...
                  </>
                ) : (
                  <>
                    <Play className="w-6 h-6 fill-current" />
                    추첨 진행하기
                  </>
                )}
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6 animate-in fade-in">
            {isProcessing && (
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 flex flex-col items-center justify-center min-h-[600px]">
                <div className="text-7xl mb-6 animate-[spin_3s_linear_infinite]">🎲</div>
                <h3 className="text-2xl font-bold text-blue-600">데이터 정합성 검증 및 추첨 중...</h3>
              </div>
            )}
            
            {results && analysis && !isProcessing && (
              <>
                <div className="bg-white rounded-xl shadow-xl border-b-2 border-slate-100 p-6 flex flex-col sm:flex-row justify-between items-center sticky top-0 z-10 gap-4">
                  <div>
                    <h2 className="text-2xl font-bold text-slate-800 flex items-center gap-3">
                      <span className="bg-blue-100 text-blue-600 w-10 h-10 rounded-full flex items-center justify-center text-xl font-bold">3</span>
                      통합 추첨 결과
                    </h2>
                    <p className="text-sm text-slate-500 font-medium mt-2">
                      전체 강좌: {Object.keys(results).length}개 | 총 신청인원: {Object.keys(parsedData!.applicants).length}명
                    </p>
                  </div>
                  <div className="flex gap-2 flex-wrap justify-end">
                    <button onClick={() => setShowSettingsModal(true)} className="bg-white px-4 py-2 rounded-lg shadow-sm border border-slate-200 text-slate-700 font-bold hover:bg-slate-50 flex items-center gap-1">
                      ⚙️ 설정 확인
                    </button>
                    <button onClick={downloadFullExcel} className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2 px-6 rounded-lg transition-all shadow-md flex items-center shrink-0 gap-2">
                      <FileDown className="w-5 h-5" />
                      통합 엑셀 저장
                    </button>
                  </div>
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
                    onClick={() => setActiveTab('duplicates')} 
                    className={cn("flex-1 py-2 text-xs font-bold rounded transition-colors", activeTab === 'duplicates' ? "bg-white shadow-sm text-slate-900" : "text-slate-600 hover:bg-slate-300/50")}
                  >
                    중복 관리
                  </button>
                  <button 
                    onClick={() => setActiveTab('unselected')} 
                    className={cn("flex-1 py-2 text-xs font-bold rounded transition-colors", activeTab === 'unselected' ? "bg-white shadow-sm text-slate-900" : "text-slate-600 hover:bg-slate-300/50")}
                  >
                    미선정 관리
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
                              <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden mb-4">
                                <div className={cn("h-full transition-all duration-500", barColor)} style={{ width: `${barWidth}%` }}></div>
                              </div>
                            </div>
                          </div>
                          <div className="flex flex-col gap-6">
                            <div>
                              <h4 className="text-sm font-bold text-emerald-800 mb-2">선정 명단</h4>
                              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-2">
                                {res.winners.map((w, idx) => (
                                  <div key={idx} className={cn("p-2 rounded flex justify-between items-start shadow-sm", w.isPriority ? "bg-blue-50 border-l-4 border-blue-500" : "bg-emerald-50 border-l-4 border-emerald-500")}>
                                    <div className="flex flex-col">
                                      <span className="text-sm text-slate-800 font-bold truncate" title={w.name}>{idx + 1}. {w.name}</span>
                                      <span className="text-[10px] text-slate-500 mt-0.5 truncate" title={w.memberId}>{w.memberId}</span>
                                    </div>
                                  </div>
                                ))}
                                {res.winners.length === 0 && <p className="text-sm text-slate-400 italic">선정자가 없습니다.</p>}
                              </div>
                            </div>
                            <div>
                              <h4 className="text-sm font-bold text-orange-800 mb-2">대기 순번</h4>
                              <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-2">
                                {res.waiting.map((w, idx) => (
                                  <div key={idx} className="bg-orange-50 border-l-4 border-orange-500 p-2 rounded shadow-sm flex flex-col justify-center">
                                    <span className="text-sm text-slate-700 font-bold truncate" title={w.name}>{idx + 1}. {w.name}</span>
                                    <span className="text-[10px] text-slate-500 mt-0.5 truncate" title={w.memberId}>{w.memberId}</span>
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

                {/* Duplicates View */}
                {activeTab === 'duplicates' && (
                  <div className="space-y-8 animate-in fade-in">
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
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-2">
                        {analysis.duplicates.length > 0 ? analysis.duplicates.map((m, idx) => (
                          <div key={idx} className="bg-red-50 border-l-4 border-red-500 p-2 rounded flex flex-col shadow-sm">
                            <div className="flex justify-between items-center mb-1">
                              <span className="text-sm font-bold text-slate-800 truncate" title={m.name}>{m.name}</span>
                              <span className="text-[10px] bg-red-100 text-red-600 px-1.5 rounded flex items-center shrink-0">{m.winCount}개</span>
                            </div>
                            <span className="text-[10px] text-slate-500 truncate" title={m.wonCourses.join(', ')}>{m.wonCourses.join(', ')}</span>
                          </div>
                        )) : (
                          <p className="col-span-full text-center text-slate-400 py-4 italic">중복 선정자가 없습니다.</p>
                        )}
                      </div>
                    </div>
                  </div>
                )}

                {/* Unselected View */}
                {activeTab === 'unselected' && (
                  <div className="space-y-8 animate-in fade-in">
                    {/* Unselected */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-8 border-l-8 border-l-slate-400">
                      <div className="mb-6 flex justify-between items-end">
                        <div>
                          <h3 className="text-xl font-bold text-slate-900 flex items-center">
                            <span className="mr-2">👤</span> 최종 미선정자 명단
                          </h3>
                          <p className="text-sm text-slate-500 mt-1 italic">단 하나의 강좌도 선정되지 못한 분들입니다.</p>
                        </div>
                        <span className="text-slate-600 font-bold text-sm">{analysis.unselected.length}명</span>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-2">
                        {analysis.unselected.length > 0 ? analysis.unselected.map((m, idx) => (
                          <div key={idx} className="bg-slate-50 border-l-4 border-slate-500 p-2 rounded flex flex-col justify-center shadow-sm">
                            <span className="text-sm font-bold text-slate-700 truncate" title={m.name}>{idx + 1}. {m.name}</span>
                            <span className="text-[10px] text-slate-500 truncate" title={m.memberId}>{m.memberId}</span>
                          </div>
                        )) : (
                          <p className="col-span-full text-center text-slate-400 py-4 italic">미선정자가 없습니다.</p>
                        )}
                      </div>
                    </div>
                  </div>
                )}
                
                <div className="pb-20"></div>
              </>
            )}
          </div>
        )}
      </div>

      {showSettingsModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-slate-100 rounded-xl shadow-xl w-full max-w-7xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="p-4 border-b border-slate-200 bg-white flex justify-between items-center">
              <h2 className="text-xl font-bold text-slate-800">설정 확인 및 수정</h2>
              <button onClick={() => setShowSettingsModal(false)} className="text-slate-400 hover:text-slate-600">
                <X className="w-6 h-6" />
              </button>
            </div>
            <div className="p-6 overflow-y-auto flex-1">
               {renderSettingsContent()}
            </div>
            <div className="p-4 border-t border-slate-200 bg-white flex justify-end gap-2">
              <button 
                onClick={() => setShowSettingsModal(false)} 
                className="px-4 py-2 bg-slate-200 text-slate-700 rounded-lg font-bold hover:bg-slate-300 transition-colors"
              >
                닫기
              </button>
              <button 
                onClick={() => {
                  runLottery();
                  setShowSettingsModal(false);
                }} 
                className="px-4 py-2 bg-blue-600 text-white rounded-lg font-bold hover:bg-blue-700 transition-colors flex items-center gap-2"
              >
                <Play className="w-4 h-4 fill-current" /> 재추첨 실행
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
