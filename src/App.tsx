/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  LogIn, 
  MapPin, 
  Clock, 
  Briefcase, 
  Calendar, 
  FileText, 
  LogOut, 
  ChevronRight, 
  AlertCircle,
  CheckCircle2,
  Lock,
  Search,
  Plus,
  Users,
  Download,
  Trash2,
  UserPlus
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format, parseISO, differenceInMinutes, startOfMonth, endOfMonth, eachDayOfInterval, isSameDay } from 'date-fns';
import { ko } from 'date-fns/locale';
import { cn } from './lib/utils';
import * as XLSX from 'xlsx';
import { 
  collection, 
  onSnapshot, 
  doc, 
  setDoc, 
  addDoc, 
  query, 
  orderBy, 
  getDocs,
  where
} from 'firebase/firestore';
import { signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { db, auth } from './firebase';

// --- Types ---

type UserRole = 'staff' | 'admin';

interface User {
  id: string; // This will be the Staff ID used for login
  name: string;
  position: string;
  role: UserRole;
  email?: string;
}

type RecordType = 'check-in' | 'check-out' | 'field-work' | 'leave';

interface AttendanceRecord {
  id: string;
  userId: string;
  type: RecordType;
  timestamp: string;
  gps: string;
  details?: {
    place?: string;
    purpose?: string;
    duration?: string;
    leaveType?: string;
  };
}

// --- Constants & Mock Data ---

const SYSTEM_NAME = "면목 스마트 근태 관리 시스템";
const CENTER_NAME = "면목종합사회복지관";
const SECURITY_TAG = "[보안 시스템 인증 데이터]";

const MOCK_USERS: User[] = [
  {
    id: 'admin',
    name: '김면목',
    position: '수석 행정 관리자',
    role: 'admin'
  },
  {
    id: 'staff01',
    name: '이행정',
    position: '사회복지사',
    role: 'staff'
  }
];

// --- Components ---

const SecurityHeader = () => (
  <div className="security-header flex items-center gap-1.5 font-mono">
    <Lock size={10} />
    {SECURITY_TAG}
  </div>
);

const LoadingOverlay = ({ message }: { message: string }) => (
  <div className="fixed inset-0 z-50 flex flex-col items-center justify-center bg-white/90 backdrop-blur-sm">
    <div className="w-12 h-12 border-4 border-neutral-200 border-t-neutral-900 rounded-full animate-spin mb-4" />
    <p className="font-medium text-neutral-600 animate-pulse">{message}</p>
  </div>
);

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [isAuthenticating, setIsAuthenticating] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'check-in' | 'check-out' | 'field-work' | 'leave' | 'summary' | 'admin'>('dashboard');
  const [records, setRecords] = useState<AttendanceRecord[]>([]);
  const [staffList, setStaffList] = useState<User[]>(MOCK_USERS);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' | 'info' } | null>(null);
  const [serverTime, setServerTime] = useState(new Date());
  const [selectedAdminUser, setSelectedAdminUser] = useState<User | null>(null);

  // Update server time locally for the UI display
  useEffect(() => {
    const timer = setInterval(() => setServerTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  // Initialize Firebase Auth and Real-time Listeners
  useEffect(() => {
    // Attempt Auth - if fails, log but don't block
    signInAnonymously(auth).catch(err => {
      console.warn("Firebase Anonymous Auth is disabled in console. This may affect write permissions if rules are strict.", err);
    });

    // 1. Listen for Staff List Changes
    const unsubStaff = onSnapshot(collection(db, 'staff'), (snapshot) => {
      const list: User[] = [];
      snapshot.forEach(doc => {
        list.push(doc.data() as User);
      });
      
      if (list.length === 0) {
        // Seed first admin if empty
        const initialAdmin: User = { id: 'admin', name: '관리자', position: '시스템관리', role: 'admin' };
        setDoc(doc(db, 'staff', 'admin'), initialAdmin);
      } else {
        setStaffList(list);
      }
    }, (err) => {
      console.error("Staff Snapshot Error:", err);
    });

    // 2. Listen for Attendance Records
    const unsubRecords = onSnapshot(query(collection(db, 'records'), orderBy('timestamp', 'desc')), (snapshot) => {
      const list: AttendanceRecord[] = [];
      snapshot.forEach(doc => {
        list.push(doc.data() as AttendanceRecord);
      });
      setRecords(list);
    }, (err) => {
      console.error("Records Snapshot Error:", err);
    });

    return () => {
      unsubStaff();
      unsubRecords();
    };
  }, []);

  // Handle Login
  const handleLogin = (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const staffId = formData.get('staffId') as string;
    
    setIsAuthenticating(true);
    
    setTimeout(() => {
      const normalizedId = staffId.toLowerCase().trim();
      const foundUser = staffList.find(u => u.id === normalizedId);
      if (foundUser) {
        setUser(foundUser);
      } else {
        alert("등록되지 않은 직원 ID입니다. (대소문자 무관)");
      }
      setIsAuthenticating(false);
    }, 1000);
  };

  const handleLogout = () => {
    setUser(null);
    setActiveTab('dashboard');
  };

  const getServerInfo = async () => {
    try {
      const res = await fetch('/api/time');
      const data = await res.json();
      return {
        timestamp: data.timestamp,
        gps: "위치정보 확인 중..." // Default placeholder
      };
    } catch (e) {
      return {
        timestamp: new Date().toISOString(),
        gps: "위치정보 확인 불가"
      };
    }
  };

  const getClientGPS = (): Promise<string> => {
    return new Promise((resolve) => {
      if (!navigator.geolocation) {
        resolve("GPS 미지원 브라우저");
        return;
      }

      navigator.geolocation.getCurrentPosition(
        async (position) => {
          const { latitude, longitude } = position.coords;
          try {
            // Optional: Simple reverse geocoding using Nominatim (free, but use with care)
            const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${latitude}&lon=${longitude}&zoom=18&addressdetails=1`, {
              headers: { 
                'Accept-Language': 'ko-KR',
                'User-Agent': 'MyeonmokAttendanceApp/1.0'
              }
            });
            const data = await response.json();
            resolve(data.display_name || `${latitude.toFixed(6)}, ${longitude.toFixed(6)}`);
          } catch (e) {
            resolve(`${latitude.toFixed(6)}, ${longitude.toFixed(6)}`);
          }
        },
        (error) => {
          console.warn("GPS Access Error:", error);
          resolve("GPS 접근 권한 거부됨");
        },
        { enableHighAccuracy: true, timeout: 5000, maximumAge: 0 }
      );
    });
  };

  const addRecord = async (type: RecordType, details?: AttendanceRecord['details']) => {
    setIsLoading(true);
    const info = await getServerInfo();
    const realGps = await getClientGPS();
    
    const newRecord: AttendanceRecord = {
      id: (typeof crypto !== 'undefined' && crypto.randomUUID) ? crypto.randomUUID() : `rec_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`,
      userId: user!.id,
      type,
      timestamp: info.timestamp,
      gps: realGps,
      details: details || {}
    };

    try {
      await setDoc(doc(db, 'records', newRecord.id), newRecord);
    } catch (e: any) {
      console.error("Failed to save record:", e);
      alert(`기록 저장 중 오류가 발생했습니다: ${e.message || '알 수 없는 오류'}\n(Firebase 연결 및 권한 확인 필요)`);
    }
    
    setIsLoading(false);
    
    let successMsg = "정상적으로 처리되었습니다.";
    if (type === 'check-in') successMsg = "정상적으로 출근 처리되었습니다.";
    if (type === 'check-out') {
      const today = format(new Date(info.timestamp), 'yyyy-MM-dd');
      const checkIn = records.find(r => r.userId === user!.id && r.type === 'check-in' && format(parseISO(r.timestamp), 'yyyy-MM-dd') === today);
      if (checkIn) {
        const diff = differenceInMinutes(parseISO(info.timestamp), parseISO(checkIn.timestamp));
        const hours = Math.floor(diff / 60);
        const mins = diff % 60;
        successMsg = `금일 총 근무 시간은 ${hours}시간 ${mins}분입니다.`;
      } else {
        successMsg = "정상적으로 퇴근 처리되었습니다.";
      }
    }

    setMessage({ text: successMsg, type: 'success' });
    setActiveTab('dashboard');
    setTimeout(() => setMessage(null), 4000);
  };

  const isWorking = useMemo(() => {
    if (!user || !records.length) return false;
    
    try {
      const now = new Date();
      const todayStr = format(now, 'yyyy-MM-dd');
      
      // Since records are ordered by timestamp desc, the first one found for the user today is the latest
      const latestTodayRecord = records.find(r => {
        if (r.userId !== user.id) return false;
        try {
          const recDate = new Date(r.timestamp);
          return format(recDate, 'yyyy-MM-dd') === todayStr;
        } catch (err) {
          return false;
        }
      });

      if (!latestTodayRecord) return false;

      // If the last action today was check-in or field-work, they are still "working"
      // If it was check-out or leave, they are not "working" in the sense of needing a check-out
      return latestTodayRecord.type === 'check-in' || latestTodayRecord.type === 'field-work';
    } catch (e) {
      console.error("Error calculating working status:", e);
      return false;
    }
  }, [records, user?.id]);

  if (!user) {
    return (
      <div className="h-screen w-screen flex items-center justify-center p-4 bg-slate-900 overflow-hidden">
        {isAuthenticating && <LoadingOverlay message="보안 서버 인증 중..." />}
        <motion.div 
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className="w-full max-w-[400px] bg-white rounded-xl shadow-2xl p-8"
        >
          <div className="flex flex-col items-center mb-8">
            <div className="w-16 h-16 bg-blue-900 rounded-lg flex items-center justify-center text-white mb-6 shadow-lg">
              <Lock size={32} />
            </div>
            <div className="text-[10px] font-bold tracking-[0.2em] text-blue-600 uppercase mb-2">Admin AI Security</div>
            <h1 className="text-xl font-bold text-center leading-tight">면목종합사회복지관<br/>스마트 근태 관리 시스템</h1>
          </div>

          <p className="text-sm font-medium text-slate-500 mb-6 text-center">
            관리자로부터 부여받은 직원 ID를 입력해 주시기 바랍니다.
          </p>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-1.5 ml-1">
                직원 고유 ID
              </label>
              <input 
                name="staffId"
                type="text" 
                required 
                placeholder="Staff ID (예: staff01)"
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded focus:outline-none focus:ring-2 focus:ring-blue-600/20 focus:border-blue-600 transition-all font-mono text-sm uppercase" 
              />
            </div>
            <button type="submit" className="w-full bg-blue-600 text-white font-bold py-4 rounded shadow-lg hover:shadow-blue-600/20 hover:bg-blue-700 transition-all flex items-center justify-center gap-3 mt-6">
              <LogIn size={20} />
              시스템 진입
            </button>
          </form>

          <footer className="mt-8 pt-6 border-t border-slate-100 flex flex-col items-center gap-4">
            <button 
              onClick={() => {
                const input = document.querySelector('input[name="staffId"]') as HTMLInputElement;
                if (input) {
                  input.value = 'admin';
                  alert("아이디에 'admin'이 입력되었습니다. 시스템 진입 버튼을 눌러주세요.");
                }
              }}
              className="text-[10px] font-bold text-slate-400 hover:text-blue-600 transition-all underline underline-offset-4"
            >
              관리자라면 이 버튼을 눌러주세요 (admin 자동입력)
            </button>
            <div className="security-badge">{SECURITY_TAG}</div>
            <p className="text-[10px] text-slate-400 leading-relaxed font-mono text-center">
              ACCESS IS RESTRICTED TO AUTHORIZED PERSONNEL ONLY. ALL ACTIVITIES ARE LOGGED AND MONITORED.
            </p>
          </footer>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex h-screen w-screen bg-slate-100 overflow-hidden font-sans text-slate-900 select-none">
      {isLoading && <LoadingOverlay message="데이터 무결성 검증 및 기록 중..." />}
      
      {/* Main Content Area */}
      <main className="flex-1 flex flex-col bg-white overflow-hidden">
        {/* Header Bar - Responsive Optimization */}
        <header className="h-auto md:h-24 flex-shrink-0 border-b border-slate-200 flex flex-col md:flex-row items-stretch md:items-center justify-between px-3 md:px-8 py-3 md:py-0 bg-white z-10 shadow-sm gap-2">
          <div className="flex items-center gap-3 w-full md:w-auto">
            <div className="bg-blue-900 p-2 md:p-3 rounded-lg text-white flex-shrink-0">
              <Lock size={16} className="md:w-5 md:h-5" />
            </div>
            <div className="flex flex-col min-w-0">
              <span className="security-badge mb-0.5 md:mb-1.5 hidden md:block text-[8px] md:text-[10px]">{SECURITY_TAG}</span>
              <h2 className="text-sm md:text-xl font-bold text-slate-800 tracking-tight leading-tight line-clamp-1 md:line-clamp-none">
                {CENTER_NAME} <span className="md:hidden">근태</span><span className="hidden md:inline">스마트 근태관리</span>
              </h2>
            </div>
          </div>
          
          <div className="flex flex-col md:flex-row items-stretch md:items-center gap-2 md:gap-8 w-full md:w-auto">
            {/* View Switcher Tabs - Mobile Compact */}
            <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 w-full md:w-auto">
              <button 
                onClick={() => setActiveTab('dashboard')}
                className={cn(
                  "flex-1 md:flex-none px-2 md:px-4 py-1.5 md:py-2 text-[10px] md:text-xs font-bold rounded-md transition-all flex items-center justify-center gap-1",
                  activeTab === 'dashboard' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                기록
              </button>
              <button 
                onClick={() => setActiveTab('summary')}
                className={cn(
                  "flex-1 md:flex-none px-2 md:px-4 py-1.5 md:py-2 text-[10px] md:text-xs font-bold rounded-md transition-all flex items-center justify-center gap-1",
                  activeTab === 'summary' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                조회
              </button>
              {user.role === 'admin' && (
                <button 
                  onClick={() => setActiveTab('admin')}
                  className={cn(
                    "flex-1 md:flex-none px-2 md:px-4 py-1.5 md:py-2 text-[10px] md:text-xs font-bold rounded-md transition-all flex items-center justify-center gap-1",
                    activeTab === 'admin' ? "bg-slate-900 text-white shadow-lg" : "bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-100"
                  )}
                >
                  <Users size={12} className={activeTab === 'admin' ? "text-blue-400" : "text-blue-600"} />
                  관리자 전용
                </button>
              )}
            </div>

            <div className="flex items-center justify-between md:justify-end gap-3 md:gap-8 px-1 md:px-0">
              <div className="flex items-center gap-2">
                <div className="h-6 md:h-10 w-[1px] bg-slate-200 hidden md:block"></div>
                <div className="text-left md:text-right">
                  <div className="mono-label text-[8px] md:text-[10px] leading-none mb-0.5">TIME</div>
                  <div className="text-xs md:text-lg font-mono font-bold text-slate-700 leading-none">{format(serverTime, 'HH:mm:ss')}</div>
                </div>
              </div>

              <div className="h-6 md:h-10 w-[1px] bg-slate-200"></div>

              <div className="flex items-center gap-2">
                <div className="text-right">
                  <div className="text-[10px] md:text-xs font-bold text-slate-900 truncate max-w-[60px] leading-none mb-0.5">{user.name}</div>
                  <div className="text-[8px] text-blue-500 font-mono leading-none">AUTH_OK</div>
                </div>
                <button 
                  onClick={handleLogout}
                  className="p-1.5 bg-slate-50 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-full transition-all border border-slate-100"
                >
                  <LogOut size={14} />
                </button>
              </div>
            </div>
          </div>
        </header>

        {/* Scrollable Content View */}
        <div className="flex-1 overflow-y-auto bg-slate-50/30 p-4 md:p-8">
          <div className="max-w-6xl mx-auto">
            <AnimatePresence mode="wait">
              {message && (
                <motion.div 
                  initial={{ opacity: 0, y: -20 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -20 }}
                  className={cn(
                    "mb-6 p-3 md:p-4 rounded bg-white shadow-sm border-l-4 flex items-center justify-between",
                    message.type === 'success' ? "border-emerald-500" : "border-rose-500"
                  )}
                >
                  <div className="flex items-center gap-2 text-slate-800">
                    {message.type === 'success' ? <CheckCircle2 className="text-emerald-500" size={18} /> : <AlertCircle className="text-rose-500" size={18} />}
                    <span className="text-xs md:text-sm font-bold">{message.text}</span>
                  </div>
                  <button onClick={() => setMessage(null)} className="text-slate-400">
                    <LogOut size={14} className="rotate-90" />
                  </button>
                </motion.div>
              )}

              {activeTab === 'dashboard' && (
                <motion.div 
                  key="dashboard"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="flex flex-col gap-4 md:gap-8"
                >
                  {/* Attendance Actions Grid */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
                    {/* Check In/Out Section */}
                    <div className="md:col-span-2 card-geometric p-4 md:p-6 flex flex-col gap-4 md:gap-6">
                      <div className="flex items-center justify-between">
                        <h3 className="text-xs md:text-sm font-bold text-slate-700 flex items-center gap-2">
                          <Clock size={16} className="text-blue-600" /> 실시간 출퇴근
                        </h3>
                        <span className={cn(
                          "px-2 py-0.5 rounded text-[8px] md:text-[10px] font-bold uppercase",
                          isWorking ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 text-slate-500"
                        )}>
                          {isWorking ? "근무 중" : "업무 종료"}
                        </span>
                      </div>
                      <div className="grid grid-cols-2 gap-3 md:gap-4">
                        <button 
                          onClick={() => addRecord('check-in')}
                          disabled={isWorking}
                          className={cn(
                            "py-4 md:py-6 rounded-lg font-bold flex flex-col items-center justify-center gap-1.5 md:gap-2 transition-all",
                            !isWorking ? "bg-emerald-600 text-white shadow-lg shadow-emerald-600/20 hover:bg-emerald-700" : "bg-slate-50 text-slate-300 border border-slate-100 cursor-not-allowed"
                          )}
                        >
                          <LogIn size={20} className="md:w-6 md:h-6" />
                          <span className="text-[11px] md:text-sm">출근 등록</span>
                        </button>
                        <button 
                          onClick={() => addRecord('check-out')}
                          disabled={!isWorking}
                          className={cn(
                            "py-4 md:py-6 rounded-lg font-bold flex flex-col items-center justify-center gap-1.5 md:gap-2 transition-all",
                            isWorking ? "bg-rose-600 text-white shadow-lg shadow-rose-600/20 hover:bg-rose-700" : "bg-slate-50 text-slate-300 border border-slate-100 cursor-not-allowed"
                          )}
                        >
                          <LogOut size={20} className="rotate-180 md:w-6 md:h-6" />
                          <span className="text-[11px] md:text-sm">퇴근 등록</span>
                        </button>
                      </div>
                    </div>

                    {/* Field Work Form */}
                    <div className="card-geometric p-4 md:p-6 flex flex-col gap-3 md:gap-4">
                      <h3 className="text-xs md:text-sm font-bold text-slate-700 flex items-center gap-2">
                        <Briefcase size={16} className="text-blue-600" /> 외근 기록
                      </h3>
                      <form onSubmit={(e) => {
                        e.preventDefault();
                        const fd = new FormData(e.currentTarget);
                        const start = fd.get('startTime') as string;
                        const end = fd.get('endTime') as string;
                        addRecord('field-work', {
                          place: fd.get('place') as string,
                          purpose: fd.get('purpose') as string,
                          duration: `${start} ~ ${end}`
                        });
                        (e.target as HTMLFormElement).reset();
                      }} className="space-y-2 md:space-y-3">
                        <input name="place" required placeholder="방문지 명칭" className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded text-[11px] md:text-xs focus:outline-none focus:border-blue-600" />
                        <input name="purpose" required placeholder="업무 목적" className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded text-[11px] md:text-xs focus:outline-none focus:border-blue-600" />
                        <div className="grid grid-cols-2 gap-2">
                          <div className="flex flex-col gap-0.5 md:gap-1">
                            <span className="text-[8px] md:text-[9px] font-bold text-slate-400 ml-1">시작</span>
                            <input name="startTime" type="time" required className="w-full px-2 md:px-3 py-1.5 md:py-2 bg-slate-50 border border-slate-200 rounded text-[10px] md:text-xs focus:outline-none focus:border-blue-600" />
                          </div>
                          <div className="flex flex-col gap-0.5 md:gap-1">
                            <span className="text-[8px] md:text-[9px] font-bold text-slate-400 ml-1">종료</span>
                            <input name="endTime" type="time" required className="w-full px-2 md:px-3 py-1.5 md:py-2 bg-slate-50 border border-slate-200 rounded text-[10px] md:text-xs focus:outline-none focus:border-blue-600" />
                          </div>
                        </div>
                        <button type="submit" className="w-full py-2 bg-slate-900 text-white text-[11px] font-bold rounded hover:bg-slate-800 transition-colors">외근 기록 저장</button>
                      </form>
                    </div>

                    {/* Leave Form */}
                    <div className="card-geometric p-4 md:p-6 flex flex-col gap-3 md:gap-4">
                      <h3 className="text-xs md:text-sm font-bold text-slate-700 flex items-center gap-2">
                        <Calendar size={16} className="text-blue-600" /> 연차 등록
                      </h3>
                      <form onSubmit={(e) => {
                        e.preventDefault();
                        const fd = new FormData(e.currentTarget);
                        const start = fd.get('leaveStartTime') as string;
                        const end = fd.get('leaveEndTime') as string;
                        addRecord('leave', {
                          leaveType: fd.get('leaveType') as string,
                          duration: `${start} ~ ${end}`
                        });
                        (e.target as HTMLFormElement).reset();
                      }} className="space-y-2 md:space-y-3">
                        <select name="leaveType" required className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded text-[11px] md:text-xs focus:outline-none focus:border-blue-600">
                          <option value="전일 연차">전일 연차</option>
                          <option value="오전 반차">오전 반차</option>
                          <option value="오후 반차">오후 반차</option>
                          <option value="시간 단위 연차">시간 단위 연차</option>
                          <option value="대체 휴가">대체 휴가</option>
                        </select>
                        <div className="grid grid-cols-2 gap-2">
                          <div className="flex flex-col gap-0.5 md:gap-1">
                            <span className="text-[8px] md:text-[9px] font-bold text-slate-400 ml-1">시작</span>
                            <input name="leaveStartTime" type="time" required className="w-full px-2 md:px-3 py-1.5 md:py-2 bg-slate-50 border border-slate-200 rounded text-[10px] md:text-xs focus:outline-none focus:border-blue-600" />
                          </div>
                          <div className="flex flex-col gap-0.5 md:gap-1">
                            <span className="text-[8px] md:text-[9px] font-bold text-slate-400 ml-1">종료</span>
                            <input name="leaveEndTime" type="time" required className="w-full px-2 md:px-3 py-1.5 md:py-2 bg-slate-50 border border-slate-200 rounded text-[10px] md:text-xs focus:outline-none focus:border-blue-600" />
                          </div>
                        </div>
                        <button type="submit" className="w-full py-2 bg-blue-600 text-white text-[11px] font-bold rounded hover:bg-blue-700 transition-colors">연차 등록 저장</button>
                      </form>
                    </div>
                  </div>

                  {/* Summary Preview (Lower Portion of Dashboard) */}
                  <div className="flex flex-col gap-3 md:gap-4 mt-2">
                    <div className="flex items-center justify-between">
                       <h3 className="text-sm md:text-md font-bold text-slate-800">최근 기록</h3>
                       <button onClick={() => setActiveTab('summary')} className="text-[10px] md:text-xs font-bold text-blue-600 hover:underline">전체보기</button>
                    </div>
                    <SummaryViewGeometric records={records} user={user} limit={5} />
                  </div>
                </motion.div>
              )}

              {activeTab === 'summary' && (
                <SummaryViewGeometric records={records} user={user} />
              )}

              {activeTab === 'admin' && user.role === 'admin' && (
                <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col gap-6 md:gap-8">
                  {/* Staff Management Section */}
                  <div className="card-geometric p-4 md:p-8">
                    <div className="flex flex-col md:flex-row md:items-center justify-between mb-6 md:mb-8 gap-4">
                      <div>
                        <h3 className="text-sm md:text-lg font-bold text-slate-800 flex items-center gap-2">
                          <Users size={16} className="text-blue-600" /> 직원 명부 관리
                        </h3>
                        <p className="text-[10px] text-slate-400 mt-1 font-mono uppercase">Id-based authentication provisioning</p>
                      </div>
                      <button 
                        onClick={() => {
                          const data = records.map(r => {
                            const employee = staffList.find(s => s.id === r.userId);
                            return {
                              '직원ID': r.userId,
                              '성명': employee?.name || 'Unknown',
                              '직위': employee?.position || '-',
                              '구분': r.type,
                              '일시': format(parseISO(r.timestamp), 'yyyy-MM-dd HH:mm:ss', { locale: ko }),
                              'GPS 정보': r.gps,
                              '상세내용': r.type === 'field-work' ? `[${r.details?.place}] ${r.details?.purpose}` : r.details?.leaveType || r.details?.duration || '-'
                            };
                          });
                          const ws = XLSX.utils.json_to_sheet(data);
                          const wb = XLSX.utils.book_new();
                          XLSX.utils.book_append_sheet(wb, ws, "전체근태기록");
                          XLSX.writeFile(wb, `전체근태기록_${format(new Date(), 'yyyyMMdd_HHmm')}.xlsx`);
                        }}
                        className="bg-emerald-600 text-white px-4 py-2 rounded font-bold text-xs hover:bg-emerald-700 transition-all flex items-center justify-center gap-2 shadow-sm"
                      >
                        <Download size={14} /> 전체 데이터 엑셀 추출
                      </button>
                    </div>

                    {/* Inline Add Staff Form */}
                    <div className="mb-8 p-4 bg-slate-50 rounded-xl border border-slate-200">
                      <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4 flex items-center gap-2">
                        <Plus size={12} /> 신규 직원 등록
                      </h4>
                      <form 
                        onSubmit={async (e) => {
                          e.preventDefault();
                          const fd = new FormData(e.currentTarget);
                          const name = fd.get('name') as string;
                          const position = fd.get('position') as string;
                          const idInput = fd.get('id') as string;
                          const id = idInput.toLowerCase().trim();
                          
                          if (staffList.some(s => s.id === id)) {
                            alert("이미 존재하는 직원 ID입니다.");
                            return;
                          }
                          
                          if (!id || id.length < 2) {
                            alert("ID는 최소 2자 이상이어야 합니다.");
                            return;
                          }
                          
                          const newStaff: User = { id, name, position, role: 'staff' };
                          
                          try {
                            await setDoc(doc(db, 'staff', id), newStaff);
                            (e.target as HTMLFormElement).reset();
                            setMessage({ text: `${name} 직원이 시스템에 등록되었습니다.`, type: 'success' });
                            setTimeout(() => setMessage(null), 3000);
                          } catch (err) {
                            console.error(err);
                            alert("직원 등록 중 오류가 발생했습니다.");
                          }
                        }}
                        className="grid grid-cols-1 md:grid-cols-4 gap-3"
                      >
                        <input name="name" required placeholder="성명" className="px-3 py-2 bg-white border border-slate-200 rounded text-xs focus:outline-none focus:border-blue-600 font-bold" />
                        <input name="position" required placeholder="직위 (예: 사회복지사)" className="px-3 py-2 bg-white border border-slate-200 rounded text-xs focus:outline-none focus:border-blue-600" />
                        <input name="id" required placeholder="로그인 ID (영어/숫자)" className="px-3 py-2 bg-white border border-slate-200 rounded text-xs focus:outline-none focus:border-blue-600 font-mono uppercase" />
                        <button type="submit" className="bg-slate-900 text-white px-4 py-2 rounded font-bold text-xs hover:bg-black transition-all flex items-center justify-center gap-2">
                          <UserPlus size={14} /> 직원 등록 완료
                        </button>
                      </form>
                    </div>

                    <div className="overflow-x-auto -mx-4 md:mx-0">
                      <table className="w-full text-left border-collapse min-w-[500px]">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200 text-[9px] md:text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                            <th className="px-4 md:px-6 py-3 md:py-4">ID</th>
                            <th className="px-4 md:px-6 py-3 md:py-4">성명</th>
                            <th className="px-4 md:px-6 py-3 md:py-4">직위</th>
                            <th className="px-4 md:px-6 py-3 md:py-4">권한</th>
                            <th className="px-4 md:px-6 py-3 md:py-4 text-right">작업</th>
                          </tr>
                        </thead>
                        <tbody className="text-[11px] md:text-xs font-medium">
                          {staffList.map((s) => (
                            <tr key={s.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors">
                              <td className="px-4 md:px-6 py-3 md:py-4 font-mono font-bold text-blue-600">{s.id}</td>
                              <td className="px-4 md:px-6 py-3 md:py-4 font-bold text-slate-800">{s.name}</td>
                              <td className="px-4 md:px-6 py-3 md:py-4 text-slate-500">{s.position}</td>
                              <td className="px-4 md:px-6 py-3 md:py-4">
                                <span className={cn(
                                  "px-2 py-0.5 rounded-[4px] text-[8px] md:text-[9px] font-bold uppercase",
                                  s.role === 'admin' ? "bg-slate-900 text-white" : "bg-slate-100 text-slate-500"
                                )}>
                                  {s.role}
                                </span>
                              </td>
                              <td className="px-4 md:px-6 py-3 md:py-4 text-right">
                                {s.id !== user.id && (
                                  <button 
                                    onClick={async () => {
                                      if (confirm(`${s.name} 직원을 삭제하시겠습니까?`)) {
                                        try {
                                          const { deleteDoc, doc } = await import('firebase/firestore');
                                          await deleteDoc(doc(db, 'staff', s.id));
                                          setMessage({ text: "직원 정보가 삭제되었습니다.", type: 'info' });
                                          setTimeout(() => setMessage(null), 3000);
                                        } catch (e) {
                                          console.error(e);
                                          alert("삭제에 실패했습니다.");
                                        }
                                      }
                                    }}
                                    className="text-rose-400 hover:text-rose-600 transition-colors"
                                  >
                                    <Trash2 size={16} />
                                  </button>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Comprehensive Records Review Section */}
                  <div className="card-geometric p-4 md:p-8">
                    <h3 className="text-sm md:text-lg font-bold text-slate-800 mb-4 md:mb-6 flex items-center gap-2">
                       <FileText size={16} className="text-blue-600" /> 전 직원 근태 모니터링
                    </h3>

                    <div className="flex gap-4 mb-6 md:mb-8">
                       <div className="flex-1 relative">
                         <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                           <Search size={14} />
                         </div>
                         <select 
                           onChange={(e) => {
                             const found = staffList.find(u => u.id === e.target.value);
                             setSelectedAdminUser(found || null);
                           }}
                           className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded text-xs md:text-sm font-bold focus:outline-none focus:border-blue-600 appearance-none"
                         >
                           <option value="">-- 검토할 직원 선택 --</option>
                           {staffList.map(s => <option key={s.id} value={s.id}>{s.name} ({s.position})</option>)}
                         </select>
                       </div>
                    </div>

                    {selectedAdminUser ? (
                      <div className="space-y-4 md:space-y-6">
                        <div className="p-3 md:p-4 bg-blue-50 border border-blue-100 rounded flex flex-col md:flex-row md:items-center justify-between gap-2">
                          <div className="text-[11px] md:text-sm font-bold text-blue-900">
                            대상: {selectedAdminUser.name} | {selectedAdminUser.position} ({selectedAdminUser.id})
                          </div>
                          <span className="text-[8px] md:text-[10px] font-bold text-blue-400 font-mono">SELECTED_FOR_AUDIT</span>
                        </div>
                        <SummaryViewGeometric records={records.filter(r => r.userId === selectedAdminUser.id)} user={selectedAdminUser} />
                      </div>
                    ) : (
                      <div className="py-12 md:py-20 text-center border-2 border-dashed border-slate-100 rounded-xl text-slate-300 font-bold italic text-xs md:text-sm">
                        직원을 선택하면 데이터가 로드됩니다.
                      </div>
                    )}
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>

        {/* Footer Branding */}
        <footer className="h-16 px-8 flex-shrink-0 bg-slate-50 border-t border-slate-200 flex justify-between items-center z-10 font-mono">
          <div className="text-[10px] text-slate-400 uppercase tracking-widest font-bold">
            © {format(new Date(), 'yyyy')} Myeonmok Social Welfare Center | Security Protocol v4.0.2
          </div>
          <div className="flex gap-3">
            <div className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse"></div>
            <div className="h-2 w-2 rounded-full bg-slate-200"></div>
            <div className="h-2 w-2 rounded-full bg-slate-200"></div>
          </div>
        </footer>
      </main>
    </div>
  );
}

// --- Sub-components (Geometric Balance Theme) ---

function NavButton({ id, label, active, onClick }: { id: string, label: string, active: boolean, onClick: () => void }) {
  return (
    <button 
      onClick={onClick}
      className={cn(
        "sidebar-btn",
        active ? "sidebar-btn-active" : "sidebar-btn-inactive"
      )}
    >
      <span className="opacity-50 text-xs font-mono">{id}</span> {label}
    </button>
  );
}

function StatCard({ label, value, subtext, active }: { label: string, value: string, subtext: string, active?: boolean }) {
  return (
    <div className="p-6 card-geometric bg-slate-50 flex flex-col gap-1 border-slate-200">
      <span className="mono-label">{label}</span>
      <div className={cn("text-2xl font-bold tracking-tight", active ? "text-emerald-600" : "text-slate-800")}>{value}</div>
      <div className="text-xs text-slate-500 font-medium">{subtext}</div>
    </div>
  );
}

function FormInput({ label, name, type = "text", placeholder }: { label: string, name: string, type?: string, placeholder?: string }) {
  return (
    <div>
      <label className="mono-label mb-2 block">{label}</label>
      <input 
        name={name} 
        type={type} 
        required 
        placeholder={placeholder}
        className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded focus:border-blue-600 focus:outline-none text-sm font-medium transition-colors" 
      />
    </div>
  );
}

function StepViewGeometric({ title, description, onConfirm, onCancel, icon }: { title: string, description: string, onConfirm: () => void, onCancel: () => void, icon: React.ReactNode }) {
  return (
    <motion.div 
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="max-w-lg mx-auto card-geometric p-10 flex flex-col items-center text-center"
    >
      <div className="w-20 h-20 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-8">
        {icon}
      </div>
      <h3 className="text-xl font-bold text-slate-800 mb-4">{title} 프로세스</h3>
      <p className="text-sm text-slate-500 leading-relaxed mb-10">
        {description}<br />
        <span className="font-mono text-[10px] text-blue-600 bg-blue-50 px-2 py-0.5 rounded mt-4 inline-block font-bold">WGS84_GPS_ACTIVE | SERVER_PORT_SYNCHRONIZED</span>
      </p>
      <div className="flex gap-4 w-full">
        <button onClick={onCancel} className="flex-1 py-4 bg-slate-100 text-slate-600 font-bold rounded transition-colors hover:bg-slate-200">취소</button>
        <button onClick={onConfirm} className="flex-1 py-4 bg-blue-600 text-white font-bold rounded shadow-lg shadow-blue-600/20 hover:bg-blue-700 transition-all">공식 기록 생성</button>
      </div>
    </motion.div>
  );
}

function SummaryViewGeometric({ records, user, limit }: { records: AttendanceRecord[], user: User, limit?: number }) {
  const currentMonth = new Date();
  const allDays = eachDayOfInterval({
    start: startOfMonth(currentMonth),
    end: endOfMonth(currentMonth)
  });

  // If limit is provided, only show the most recent days that have records or are today/past
  const displayedDays = limit ? allDays.filter(d => d <= new Date()).reverse().slice(0, limit) : allDays;

  const totalMinutes = useMemo(() => {
    let total = 0;
    allDays.forEach(day => {
      const dayRecords = records.filter(r => isSameDay(parseISO(r.timestamp), day));
      const cin = dayRecords.find(r => r.type === 'check-in');
      const cout = dayRecords.find(r => r.type === 'check-out');
      if (cin && cout) {
        total += differenceInMinutes(parseISO(cout.timestamp), parseISO(cin.timestamp));
      }
    });
    return total;
  }, [records, allDays]);

  const handleExportExcel = () => {
    const data = allDays.map(day => {
      const dayRecords = records.filter(r => isSameDay(parseISO(r.timestamp), day));
      const cin = dayRecords.find(r => r.type === 'check-in');
      const cout = dayRecords.find(r => r.type === 'check-out');
      const fws = dayRecords.filter(r => r.type === 'field-work').map(f => `[${f.details?.place}] ${f.details?.purpose} (${f.details?.duration})`).join('\n');
      const lv = dayRecords.find(r => r.type === 'leave');

      return {
        '날짜': format(day, 'yyyy-MM-dd (EE)', { locale: ko }),
        '출근 시각': cin ? format(parseISO(cin.timestamp), 'HH:mm:ss') : '-',
        '퇴근 시각': cout ? format(parseISO(cout.timestamp), 'HH:mm:ss') : '-',
        '외근 기록': fws || '-',
        '연차 여부': lv ? lv.details?.leaveType : '-',
        '비고': ''
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '근무기록');
    XLSX.writeFile(workbook, `근무기록_${user.name}_${format(currentMonth, 'yyyyMM')}.xlsx`);
  };

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col gap-6">
      <div className="card-geometric border-slate-200 shadow-lg">
        {!limit && (
          <div className="bg-slate-50 px-8 py-5 border-b border-slate-200 flex justify-between items-center">
            <h3 className="text-sm font-bold text-slate-700">{format(currentMonth, 'yyyy년 MM월')} 근무 요약 현황</h3>
            <div className="flex items-center gap-4">
               <button 
                 onClick={handleExportExcel}
                 className="flex items-center gap-2 text-[10px] font-bold text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded hover:bg-emerald-100 transition-all"
               >
                 <Download size={14} /> EXCEL 다운로드
               </button>
               <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">KST / WGS84 GPS</span>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[800px]">
             <thead>
              <tr className="bg-slate-50 text-[10px] font-bold text-slate-400 uppercase tracking-wider border-b border-slate-200">
                <th className="px-6 py-4 border-r border-slate-100">날짜</th>
                <th className="px-6 py-4 border-r border-slate-100">출근 (GPS)</th>
                <th className="px-6 py-4 border-r border-slate-100">퇴근 (GPS)</th>
                <th className="px-6 py-4 border-r border-slate-100">외근 (장소/목적)</th>
                <th className="px-6 py-4">연차</th>
              </tr>
            </thead>
            <tbody className="text-xs text-slate-600 font-mono">
              {displayedDays.map((day, idx) => {
                const dayRecords = records.filter(r => isSameDay(parseISO(r.timestamp), day));
                const cin = dayRecords.find(r => r.type === 'check-in');
                const cout = dayRecords.find(r => r.type === 'check-out');
                const fws = dayRecords.filter(r => r.type === 'field-work');
                const lv = dayRecords.find(r => r.type === 'leave');
                const isToday = isSameDay(day, new Date());

                return (
                  <tr key={idx} className={cn("border-b border-slate-100 transition-colors", isToday ? "bg-blue-50/30" : "hover:bg-slate-50/50")}>
                    <td className="px-6 py-4 border-r border-slate-100 font-sans font-bold text-slate-900 whitespace-nowrap">
                       {format(day, 'MM-dd')} ({format(day, 'E', { locale: ko })})
                    </td>
                    <td className="px-6 py-4 border-r border-slate-100">
                      {cin ? (
                        <div className="flex flex-col">
                          <span className="text-slate-900 font-bold">{format(parseISO(cin.timestamp), 'HH:mm:ss')}</span>
                          <span className="text-[9px] opacity-40 uppercase truncate max-w-[120px]">{cin.gps}</span>
                        </div>
                      ) : <span className="text-slate-200 font-bold">-</span>}
                    </td>
                    <td className="px-6 py-4 border-r border-slate-100">
                      {cout ? (
                        <div className="flex flex-col">
                          <span className="text-slate-900 font-bold">{format(parseISO(cout.timestamp), 'HH:mm:ss')}</span>
                          <span className="text-[9px] opacity-40 uppercase truncate max-w-[120px]">{cout.gps}</span>
                        </div>
                      ) : (isToday && cin ? <span className="text-slate-400 italic">진행 중...</span> : <span className="text-slate-200 font-bold">-</span>)}
                    </td>
                    <td className="px-6 py-4 border-r border-slate-100 font-sans">
                      {fws.length > 0 ? fws.map(fw => (
                        <div key={fw.id} className="text-[11px] leading-tight mb-1">
                          <span className="text-blue-700 font-bold">[{fw.details?.place}]</span> {fw.details?.purpose}
                        </div>
                      )) : <span className="text-slate-300">-</span>}
                    </td>
                    <td className="px-6 py-4 font-sans">
                      {lv ? (
                        <span className="px-2 py-1 bg-rose-50 text-rose-700 font-bold rounded text-[10px] italic">
                          {lv.details?.leaveType}
                        </span>
                      ) : <span className="text-slate-300">-</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        {!limit && (
          <div className="bg-slate-900 px-8 py-4 flex justify-between items-center text-white">
            <span className="text-[10px] font-bold uppercase tracking-widest opacity-60 font-mono">[해당 월 총 근무 시간 합계]</span>
            <span className="text-xl font-mono font-bold">{Math.floor(totalMinutes / 60)}시간 {totalMinutes % 60}분</span>
          </div>
        )}
      </div>
    </motion.div>
  );
}
