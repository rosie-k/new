/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from "react";
import {
  Calendar,
  Download,
  Upload,
  Plus,
  Trash2,
  Edit2,
  FileSpreadsheet,
  AlertTriangle,
  CheckCircle,
  Info,
  Clock,
  ToggleLeft,
  ToggleRight,
  Moon,
  Loader2,
  Tag,
  Zap
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { Booking, TARGET_FACILITIES, TargetFacility } from "./types";
import {
  parseUploadedExcelInBrowser,
  generateTimetableExcelBlob,
  generateMockTemplateBlob,
  formatPhoneNumber,
  parseLocalDate,
  computeExceededBookings,
  getDayOfWeekKorean,
  isLightingFeeApplied,
  TimetableDay,
} from "./excelGenerator";

export default function App() {
  // Main state: reservation list
  const [bookings, setBookings] = useState<Booking[]>([]);
  const [selectedDate, setSelectedDate] = useState<string>("");
  const [activeTabDate, setActiveTabDate] = useState<string>("");
  const [showAllHoursFriday, setShowAllHoursFriday] = useState<boolean>(false);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [successMessage, setSuccessMessage] = useState<string>("");

  // Search & Filter
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [filterFacility, setFilterFacility] = useState<string>("ALL");

  // Form states for manual additions/edits
  const [isFormOpen, setIsFormOpen] = useState<boolean>(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [formFacility, setFormFacility] = useState<TargetFacility>("테니스장A");
  const [formBooker, setFormBooker] = useState<string>("");
  const [formGroup, setFormGroup] = useState<string>("");
  const [formDate, setFormDate] = useState<string>("");
  const [formStartTime, setFormStartTime] = useState<string>("09:00");
  const [formEndTime, setFormEndTime] = useState<string>("11:00");
  const [formPhone, setFormPhone] = useState<string>("");
  const [formStatus, setFormStatus] = useState<string>("예약완료");

  const formatDate = (d: Date) => {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  };

  const formatMMDD = (dateStr: string): string => {
    if (!dateStr) return "";
    const parts = dateStr.split("-");
    if (parts.length === 3) {
      return `${parts[1]}.${parts[2]}.`;
    }
    return "";
  };

  // Helper messages
  const showSuccess = (msg: string) => {
    setSuccessMessage(msg);
    setTimeout(() => setSuccessMessage(""), 4000);
  };

  const showError = (msg: string) => {
    setErrorMessage(msg);
    setTimeout(() => setErrorMessage(""), 4000);
  };

  // Standard Friday/Saturday/Sunday calculation based on selectedDate (Friday reference)
  const weekDates = useMemo(() => {
    if (!selectedDate) return { Friday: "", Saturday: "", Sunday: "" };
    const fri = parseLocalDate(selectedDate);
    const sat = new Date(fri);
    sat.setDate(fri.getDate() + 1);
    const sun = new Date(fri);
    sun.setDate(fri.getDate() + 2);

    return {
      Friday: formatDate(fri),
      Saturday: formatDate(sat),
      Sunday: formatDate(sun),
    };
  }, [selectedDate]);

  // Compute all available tab dates (Friday, Saturday, Sunday, + any holidays in bookings)
  const availableDays: TimetableDay[] = useMemo(() => {
    if (!weekDates.Friday) return [];
    const standard: TimetableDay[] = [
      { name: "금요일", date: weekDates.Friday },
      { name: "토요일", date: weekDates.Saturday },
      { name: "일요일", date: weekDates.Sunday },
    ];

    const stdSet = new Set<string>([weekDates.Friday, weekDates.Saturday, weekDates.Sunday]);
    const extraDates: string[] = Array.from(
      new Set<string>(
        bookings
          .map((b) => (b.date || "").trim())
          .filter((d): d is string => Boolean(d && !stdSet.has(d)))
      )
    ).sort();

    const holidays: TimetableDay[] = extraDates.map((dStr) => {
      const dayName = getDayOfWeekKorean(dStr);
      return {
        name: `공휴일(${dayName})`,
        date: dStr,
        isHoliday: true,
      };
    });

    return [...standard, ...holidays];
  }, [weekDates, bookings]);

  // Set default selected date
  useEffect(() => {
    const today = parseLocalDate("2026-06-28");
    const day = today.getDay();
    const distanceToFriday = 5 - day;
    const fri = new Date(today);
    fri.setDate(today.getDate() + distanceToFriday);
    const formattedFri = formatDate(fri);
    setSelectedDate(formattedFri);
    setActiveTabDate(formattedFri);
  }, []);

  // Update active tab date if it is empty or invalid
  useEffect(() => {
    if (availableDays.length > 0 && (!activeTabDate || !availableDays.some((d) => d.date === activeTabDate))) {
      setActiveTabDate(availableDays[0].date);
    }
  }, [availableDays, activeTabDate]);

  const autoSelectWeekFromBookings = (loadedBookings: Booking[]) => {
    if (!loadedBookings || loadedBookings.length === 0) return;
    const bookingWithDate = loadedBookings.find((b) => b.date && b.date.trim() !== "");
    if (!bookingWithDate) return;

    const d = parseLocalDate(bookingWithDate.date);
    if (isNaN(d.getTime())) return;

    const day = d.getDay();
    let offset = 0;
    if (day === 0) offset = -2;
    else if (day === 6) offset = -1;
    else offset = 5 - day;

    const fri = new Date(d);
    fri.setDate(d.getDate() + offset);
    setSelectedDate(formatDate(fri));
    setActiveTabDate(bookingWithDate.date);
  };

  // Upload custom Excel file (Pure Browser JavaScript execution)
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setErrorMessage("");

    try {
      const arrayBuffer = await file.arrayBuffer();
      const parsedBookings = await parseUploadedExcelInBrowser(arrayBuffer);

      if (parsedBookings.length === 0) {
        throw new Error("엑셀 파일에서 유효한 예약 데이터를 찾을 수 없습니다.");
      }

      setBookings(parsedBookings);
      autoSelectWeekFromBookings(parsedBookings);
      showSuccess(`[브라우저 파싱 완료] 총 ${parsedBookings.length}개의 예약을 성공적으로 불러왔습니다.`);
    } catch (err: any) {
      showError(err.message || "엑셀 파일 변환에 실패했습니다.");
    } finally {
      setIsUploading(false);
      e.target.value = "";
    }
  };

  // Download beautiful styled Excel (Pure Browser JavaScript execution - No Server Required)
  const handleExportExcel = async () => {
    if (bookings.length === 0) {
      showError("내보낼 예약 데이터가 없습니다. 먼저 예약 데이터를 추가하거나 수집해 주세요.");
      return;
    }

    setIsExporting(true);
    try {
      const { blob, fileName } = await generateTimetableExcelBlob(bookings, weekDates, availableDays);

      const blobUrl = URL.createObjectURL(blob);
      const downloadLink = document.createElement("a");
      downloadLink.href = blobUrl;
      downloadLink.download = fileName;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
      URL.revokeObjectURL(blobUrl);

      showSuccess(`[브라우저 즉시 생성] '${fileName}' 파일이 다운로드되었습니다.`);
    } catch (err: any) {
      showError(err.message || "엑셀 시간표 파일 생성 실패");
    } finally {
      setIsExporting(false);
    }
  };

  // Download raw mock template (Pure browser generation)
  const handleDownloadTemplate = async () => {
    try {
      const blob = await generateMockTemplateBlob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "gongyoonuri_raw_export.xlsx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      showError("샘플 파일 생성 실패: " + err.message);
    }
  };

  // Manual bookings operations
  const handleOpenAddForm = () => {
    setEditingId(null);
    setFormFacility("테니스장A");
    setFormBooker("");
    setFormGroup("");
    setFormDate(activeTabDate || weekDates.Friday);
    setFormStartTime("09:00");
    setFormEndTime("11:00");
    setFormPhone("");
    setFormStatus("예약완료");
    setIsFormOpen(true);
  };

  const handleOpenEditForm = (b: Booking) => {
    setEditingId(b.id);
    setFormFacility(b.resolvedFacility as TargetFacility);
    setFormBooker(b.bookerName);
    setFormGroup(b.groupName || "");
    setFormDate(b.date || activeTabDate);
    setFormStartTime(b.startTime);
    setFormEndTime(b.endTime);
    setFormPhone(b.phoneNumber || "");
    setFormStatus(b.status || "예약완료");
    setIsFormOpen(true);
  };

  const handleSaveForm = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formBooker.trim()) {
      showError("예약자명(또는 단체명)을 반드시 입력해주세요.");
      return;
    }

    const startHour = parseInt(formStartTime.split(":")[0], 10);
    const endHour = parseInt(formEndTime.split(":")[0], 10);
    const isNightHighlighted = isLightingFeeApplied(formDate, formStartTime, formEndTime);

    if (startHour >= endHour) {
      showError("종료 시간은 시작 시간보다 늦어야 합니다.");
      return;
    }

    if (editingId) {
      setBookings((prev) =>
        prev.map((b) =>
          b.id === editingId
            ? {
                ...b,
                resolvedFacility: formFacility,
                bookerName: formBooker,
                groupName: formGroup,
                date: formDate,
                startTime: formStartTime,
                endTime: formEndTime,
                phoneNumber: formatPhoneNumber(formPhone),
                status: formStatus,
                isNightHighlighted,
              }
            : b
        )
      );
      showSuccess("예약 정보를 성공적으로 수정하였습니다.");
    } else {
      const newBooking: Booking = {
        id: `manual-${Date.now()}`,
        bookingNo: `M-${Date.now().toString().slice(-6)}`,
        facilityName: formFacility,
        resolvedFacility: formFacility,
        bookerName: formBooker,
        groupName: formGroup,
        date: formDate,
        startTime: formStartTime,
        endTime: formEndTime,
        phoneNumber: formatPhoneNumber(formPhone),
        status: formStatus,
        isNightHighlighted,
      };
      setBookings((prev) => [...prev, newBooking]);
      showSuccess("신규 예약을 시간표에 수동 추가하였습니다.");
    }

    setIsFormOpen(false);
  };

  const handleDeleteBooking = (id: string) => {
    if (window.confirm("선택한 예약 건을 목록에서 정말 삭제하시겠습니까?")) {
      setBookings((prev) => prev.filter((b) => b.id !== id));
      showSuccess("예약이 삭제되었습니다.");
    }
  };

  // Helper lists for form
  const hoursOptions = Array.from({ length: 18 }, (_, i) => {
    const h = String(6 + i).padStart(2, "0");
    return `${h}:00`;
  });

  const processedBookings = useMemo(() => {
    return computeExceededBookings(bookings);
  }, [bookings]);

  const exceededBookings = processedBookings.filter((b) => b.isExceeded);

  // Filter bookings based on UI search/filters
  const getFilteredBookings = () => {
    let list = processedBookings;
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      list = list.filter(
        (b) =>
          b.bookerName.toLowerCase().includes(q) ||
          (b.groupName && b.groupName.toLowerCase().includes(q)) ||
          (b.bookingNo && b.bookingNo.toLowerCase().includes(q)) ||
          b.facilityName.toLowerCase().includes(q)
      );
    }
    if (filterFacility !== "ALL") {
      list = list.filter((b) => b.resolvedFacility === filterFacility);
    }
    return list;
  };

  const currentActiveDayObj = availableDays.find((d) => d.date === activeTabDate) || availableDays[0];
  const isCurrentTabFriday = currentActiveDayObj?.name === "금요일" && !currentActiveDayObj.isHoliday;
  const isCurrentTabHoliday = currentActiveDayObj?.isHoliday || false;

  const activeBookings = getFilteredBookings().filter((b) => {
    if (!b.date) return false;
    if (b.date === activeTabDate) return true;
    if (isCurrentTabHoliday) return false;
    const bDate = parseLocalDate(b.date);
    const targetDayOfWeek = currentActiveDayObj?.name === "금요일" ? 5 : currentActiveDayObj?.name === "토요일" ? 6 : 0;
    return bDate.getDay() === targetDayOfWeek;
  });

  // Time matrix construction for grid render
  const renderTimeGrid = () => {
    let hours: number[] = [];
    if (isCurrentTabFriday) {
      hours = showAllHoursFriday
        ? Array.from({ length: 16 }, (_, i) => 6 + i) // 06:00 to 22:00
        : [6, 7, 8, 18, 19, 20, 21]; // Compact: 06~09, 18~22
    } else {
      hours = Array.from({ length: 13 }, (_, i) => 7 + i); // 07:00 to 20:00
    }

    const spannedCells: Record<string, boolean> = {};

    return (
      <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-sm" id="schedule-timetable-container">
        <table className="w-full border-collapse text-left font-sans">
          <thead>
            <tr className="bg-slate-50 border-b border-slate-200">
              <th className="p-3 text-xs font-semibold text-slate-500 uppercase tracking-wider text-center border-r border-slate-200 w-32 font-mono">
                시간대
              </th>
              {TARGET_FACILITIES.map((fac) => (
                <th key={fac} className="p-3 text-xs font-semibold text-slate-700 text-center border-r border-slate-200 min-w-36">
                  {fac === "복합운동장" ? (
                    <span className="bg-sky-50 text-sky-700 px-2.5 py-1 rounded-full font-medium">{fac}</span>
                  ) : fac.startsWith("테니스") ? (
                    <span className="bg-emerald-50 text-emerald-700 px-2.5 py-1 rounded-full font-medium">{fac}</span>
                  ) : (
                    <span className="bg-orange-50 text-orange-700 px-2.5 py-1 rounded-full font-medium">{fac}</span>
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {hours.map((hour) => {
              const startStr = String(hour).padStart(2, "0") + ":00";
              const endStr = String(hour + 1).padStart(2, "0") + ":00";
              const isLightingHour = isLightingFeeApplied(activeTabDate, startStr, endStr);

              return (
                <tr
                  key={hour}
                  className={`border-b border-slate-100 hover:bg-slate-50/40 transition-colors ${
                    isLightingHour ? "bg-amber-50/20" : ""
                  }`}
                >
                  <td
                    className={`p-2.5 text-center font-mono text-xs font-medium border-r border-slate-200 ${
                      isLightingHour ? "text-amber-800 bg-amber-100/50 font-bold" : "text-slate-500 bg-slate-50/50"
                    }`}
                  >
                    <div className="flex items-center justify-center gap-1">
                      {isLightingHour ? (
                        <Zap className="w-3.5 h-3.5 text-amber-600 fill-amber-300" />
                      ) : (
                        <Clock className="w-3.5 h-3.5 text-slate-400" />
                      )}
                      <span>
                        {startStr}~{endStr}
                      </span>
                    </div>
                  </td>

                  {TARGET_FACILITIES.map((fac) => {
                    const spanKey = `${fac}-${hour}`;
                    if (spannedCells[spanKey]) return null;

                    const cellBookings = activeBookings.filter((b) => {
                      if (b.resolvedFacility !== fac) return false;
                      const bStart = parseInt(b.startTime.split(":")[0], 10);
                      const bEnd = parseInt(b.endTime.split(":")[0], 10);
                      return hour >= bStart && hour < bEnd;
                    });

                    if (cellBookings.length > 0) {
                      const earliestStart = Math.min(...cellBookings.map((b) => parseInt(b.startTime.split(":")[0], 10)));
                      const latestEnd = Math.max(...cellBookings.map((b) => parseInt(b.endTime.split(":")[0], 10)));

                      const visibleHoursSpan = hours.filter((h) => h >= earliestStart && h < latestEnd);
                      const rowSpan = visibleHoursSpan.length || 1;

                      visibleHoursSpan.forEach((h) => {
                        spannedCells[`${fac}-${h}`] = true;
                      });

                      const hasConflict = cellBookings.length > 1;
                      const primaryBooking = cellBookings[0];
                      const isNight = cellBookings.some((b) => b.isNightHighlighted || isLightingFeeApplied(b.date, b.startTime, b.endTime));
                      const isExceeded = cellBookings.some((b) => b.isExceeded);

                      return (
                        <td
                          key={fac}
                          rowSpan={rowSpan}
                          className="p-2 text-center border-r border-slate-200 transition-all align-middle select-none relative group"
                        >
                          <div
                            className={`p-3 rounded-lg border flex flex-col items-center justify-center gap-1.5 shadow-sm transition-all h-full min-h-[50px] cursor-pointer hover:shadow-md
                              ${
                                isExceeded
                                  ? "bg-pink-300 border-pink-400 text-pink-950 font-bold ring-4 ring-pink-500/20 shadow"
                                  : isNight
                                  ? "bg-yellow-300 border-yellow-400 text-black font-semibold ring-2 ring-yellow-400/20"
                                  : "bg-slate-50 border-slate-200 text-slate-800"
                              }
                            `}
                            onClick={() => handleOpenEditForm(primaryBooking)}
                          >
                            <div className="absolute top-1.5 right-1.5 hidden group-hover:flex items-center gap-1 bg-white/95 px-1 py-0.5 rounded border border-slate-200 shadow-sm">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleOpenEditForm(primaryBooking);
                                }}
                                className="p-1 hover:text-sky-600 text-slate-400 transition-colors"
                                title="수정"
                              >
                                <Edit2 className="w-3 h-3" />
                              </button>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteBooking(primaryBooking.id);
                                }}
                                className="p-1 hover:text-rose-600 text-slate-400 transition-colors"
                                title="삭제"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>

                            <div className="leading-snug tracking-tight my-1 flex flex-col items-center justify-center w-full">
                              {hasConflict ? (
                                <div className="space-y-1.5 w-full">
                                  {cellBookings.map((cb) => (
                                    <div
                                      key={cb.id}
                                      className="border-b border-slate-200/60 last:border-0 pb-1.5 last:pb-0 pt-1 last:pt-0 flex flex-col items-center justify-center"
                                    >
                                      <div
                                        className={`text-xs font-bold text-slate-900 ${
                                          cb.isExceeded ? "text-rose-900 underline decoration-rose-500 font-extrabold" : ""
                                        }`}
                                      >
                                        {cb.bookerName}
                                      </div>
                                      {cb.phoneNumber && (
                                        <div className="text-[9px] text-slate-500 font-mono mt-0.5">
                                          {formatPhoneNumber(cb.phoneNumber)}
                                        </div>
                                      )}
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                <div className="flex flex-col items-center justify-center gap-0.5 w-full">
                                  <span
                                    className={`text-[13px] font-bold text-slate-900 block ${
                                      primaryBooking.isExceeded ? "text-rose-900 underline decoration-rose-500 font-extrabold" : ""
                                    }`}
                                  >
                                    {primaryBooking.bookerName}
                                  </span>
                                  {primaryBooking.phoneNumber && (
                                    <span className="text-[9px] text-slate-500 font-mono block">
                                      {formatPhoneNumber(primaryBooking.phoneNumber)}
                                    </span>
                                  )}
                                </div>
                              )}
                            </div>
                          </div>
                        </td>
                      );
                    }

                    return (
                      <td
                        key={fac}
                        className="p-2.5 border-r border-slate-200 min-h-[50px] align-middle text-center text-slate-300 text-xs hover:bg-slate-50/50 cursor-pointer"
                        onClick={handleOpenAddForm}
                      >
                        <span className="opacity-0 hover:opacity-100 flex items-center justify-center gap-1 text-slate-400">
                          <Plus className="w-3.5 h-3.5" /> 추가
                        </span>
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-slate-50/50 text-slate-800 flex flex-col antialiased">
      {/* Top Notification Bar */}
      <AnimatePresence>
        {successMessage && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="fixed top-4 left-1/2 -translate-x-1/2 z-50 bg-emerald-600 text-white px-6 py-3 rounded-full shadow-lg flex items-center gap-2 font-medium text-sm"
          >
            <CheckCircle className="w-4 h-4" />
            <span>{successMessage}</span>
          </motion.div>
        )}
        {errorMessage && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            className="fixed top-4 left-1/2 -translate-x-1/2 z-50 bg-rose-600 text-white px-6 py-3 rounded-full shadow-lg flex items-center gap-2 font-medium text-sm"
          >
            <AlertTriangle className="w-4 h-4" />
            <span>{errorMessage}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Modern Main Header */}
      <header className="bg-white border-b border-slate-200/80 sticky top-0 z-40 px-6 py-4 shadow-sm/5%">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-gradient-to-tr from-sky-600 to-indigo-700 text-white p-2.5 rounded-xl shadow-md shadow-sky-600/15">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-bold tracking-tight text-slate-900 font-display">
                  공유누리 예약 관리 도구
                </h1>
                <span className="bg-indigo-50 text-indigo-700 text-[10px] font-bold px-2 py-0.5 rounded-full border border-indigo-200/45">
                  100% 브라우저 동작
                </span>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                공유누리 주말/공휴일 예약 데이터 변환 및 요일별 시각화 엑셀 자동 제작 솔루션
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3 self-end md:self-auto">
            {/* Week & Date Selectors */}
            <div className="flex items-center gap-2 bg-slate-100/90 p-1.5 rounded-xl border border-slate-200/50">
              <Calendar className="w-4 h-4 text-slate-500 ml-2" />
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="bg-transparent border-0 text-xs font-semibold text-slate-700 focus:ring-0 cursor-pointer pr-2 outline-none"
                title="기준 금요일 선택"
              />
              <span className="text-xs bg-white text-slate-600 font-bold px-2.5 py-1 rounded-lg border border-slate-200 shadow-sm font-mono">
                {weekDates.Friday ? `${weekDates.Friday.slice(5)} ~ ${weekDates.Sunday.slice(5)}` : ""}
              </span>
            </div>

            {/* Download Timetable Excel */}
            <button
              onClick={handleExportExcel}
              disabled={isExporting || bookings.length === 0}
              className="flex items-center gap-1.5 bg-slate-950 text-white hover:bg-slate-900 px-4 py-2.5 rounded-xl font-semibold text-xs transition-all shadow-md shadow-slate-950/10 active:scale-95 border border-slate-900 disabled:opacity-50 disabled:pointer-events-none"
            >
              {isExporting ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Download className="w-3.5 h-3.5" />}
              시각화 엑셀 저장
            </button>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl mx-auto w-full px-6 py-6 grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* LEFT COLUMN: UPLOAD & SEASONAL LIGHTING RULE GUIDE (col-span-4) */}
        <div className="lg:col-span-4 space-y-6">
          {/* STEP 1: EXCEL UPLOAD & TEMPLATE DOWNLOAD */}
          <div className="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-sm space-y-4">
            <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
              <div className="bg-emerald-50 text-emerald-600 p-1.5 rounded-lg">
                <Upload className="w-4 h-4" />
              </div>
              <h2 className="text-sm font-bold text-slate-800 font-display">1단계: 원본 엑셀 업로드</h2>
            </div>

            <p className="text-xs text-slate-500 leading-relaxed">
              공유누리에서 내려받은 원본 엑셀 파일(금/토/일 및 공휴일 포함)을 여기에 끌어놓거나 클릭하여 즉시 시각화 시간표로 변환해 보세요.
            </p>

            <div className="border-2 border-dashed border-slate-200 hover:border-sky-500 hover:bg-slate-50/50 rounded-xl p-5 text-center transition-all relative">
              <input
                type="file"
                accept=".xlsx"
                onChange={handleFileUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                disabled={isUploading}
              />
              <div className="flex flex-col items-center gap-2">
                {isUploading ? (
                  <Loader2 className="w-8 h-8 text-sky-500 animate-spin" />
                ) : (
                  <FileSpreadsheet className="w-8 h-8 text-slate-400" />
                )}
                <div className="text-xs font-semibold text-slate-600">
                  {isUploading ? "파일 파싱 및 분석 중..." : "엑셀 파일(.xlsx) 드래그 또는 업로드"}
                </div>
                <div className="text-[10px] text-slate-400">
                  공유누리 예약목록 다운로드 파일 전용 (공휴일 자동 감지)
                </div>
              </div>
            </div>

            {/* Test Sample Template */}
            <div className="bg-slate-50 rounded-xl p-3 border border-slate-150 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Info className="w-4 h-4 text-sky-600" />
                <div>
                  <div className="text-xs font-bold text-slate-700">샘플 예약 파일이 없으신가요?</div>
                  <div className="text-[10px] text-slate-500">주말+공휴일 포함 모의 데이터 다운로드</div>
                </div>
              </div>
              <button
                onClick={handleDownloadTemplate}
                className="flex items-center gap-1 bg-white hover:bg-slate-100 border border-slate-200 text-[10px] font-bold text-slate-700 px-2.5 py-1.5 rounded-lg shadow-sm"
              >
                <Download className="w-3 h-3 text-slate-500" /> 샘플 파일
              </button>
            </div>
          </div>

          {/* SEASONAL LIGHTING FEE RULES CARD */}
          <div className="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-sm space-y-3.5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="bg-amber-50 p-1.5 rounded-lg text-amber-600">
                  <Zap className="w-4 h-4" />
                </div>
                <h3 className="text-xs font-bold text-slate-800 font-display">월별 조명사용료 부과 기준</h3>
              </div>
              <span className="text-[10px] bg-amber-100 text-amber-800 font-bold px-2 py-0.5 rounded-full border border-amber-200">
                자동 감지
              </span>
            </div>

            <div className="space-y-2 text-xs">
              <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-100 space-y-1">
                <div className="font-bold text-slate-800 flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                  저녁/야간 조명료
                </div>
                <ul className="text-[11px] text-slate-600 space-y-0.5 pl-3 list-disc">
                  <li><strong>3월~4월, 9월~10월:</strong> 18:00 ~ 22:00</li>
                  <li><strong>5월~8월:</strong> 19:00 ~ 22:00</li>
                  <li><strong>11월~2월:</strong> 18:00 ~ 22:00</li>
                </ul>
              </div>

              <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-100 space-y-1">
                <div className="font-bold text-slate-800 flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-sky-500"></span>
                  새벽 조명료
                </div>
                <ul className="text-[11px] text-slate-600 space-y-0.5 pl-3 list-disc">
                  <li><strong>10월~3월:</strong> 06:00 ~ 08:00 (부과)</li>
                  <li><strong>4월~9월:</strong> 새벽 조명료 미부과</li>
                </ul>
              </div>
            </div>

            <p className="text-[10px] text-slate-400 leading-relaxed">
              * 예약 일자의 월을 자동 인식하여 조명료 부과 시간대에 해당하는 예약은 노란색(#FFFF00)으로 하이라이트됩니다.
            </p>
          </div>
        </div>

        {/* RIGHT COLUMN: TIMETABLE PREVIEW & MANAGEMENT (col-span-8) */}
        <div className="lg:col-span-8 space-y-6">
          {/* CONTROL BAR: TAB BUTTONS, SEARCH, & FILTERS */}
          <div className="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-sm space-y-4">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              {/* Day Tabs (Dynamic list with Holidays) */}
              <div className="flex flex-wrap bg-slate-100 p-1 rounded-xl border border-slate-200/40 gap-1">
                {availableDays.map((day) => {
                  const isActive = activeTabDate === day.date;
                  return (
                    <button
                      key={day.date}
                      onClick={() => setActiveTabDate(day.date)}
                      className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                        isActive
                          ? day.isHoliday
                            ? "bg-rose-600 text-white shadow-sm"
                            : "bg-white text-slate-950 shadow-sm"
                          : day.isHoliday
                          ? "text-rose-700 hover:bg-rose-50"
                          : "text-slate-500 hover:text-slate-800"
                      }`}
                    >
                      {day.isHoliday && <Tag className="w-3 h-3 text-current" />}
                      <span>{day.name}</span>
                      <span className="font-mono text-[10px] opacity-80">({formatMMDD(day.date)})</span>
                    </button>
                  );
                })}
              </div>

              {/* Utility Toggles */}
              <div className="flex items-center gap-3">
                {isCurrentTabFriday && (
                  <button
                    onClick={() => setShowAllHoursFriday(!showAllHoursFriday)}
                    className="flex items-center gap-2 bg-slate-50 hover:bg-slate-100 text-slate-600 px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-semibold"
                  >
                    {showAllHoursFriday ? "금요 핵심시간 축소" : "금요 전체시간 보기"}
                    {showAllHoursFriday ? <ToggleRight className="w-5 h-5 text-sky-600" /> : <ToggleLeft className="w-5 h-5 text-slate-400" />}
                  </button>
                )}

                <button
                  onClick={handleOpenAddForm}
                  className="flex items-center gap-1 bg-indigo-600 hover:bg-indigo-500 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all"
                >
                  <Plus className="w-3.5 h-3.5" /> 수동 추가
                </button>
              </div>
            </div>

            {/* Sub-Filters Row */}
            <div className="flex flex-col md:flex-row items-center gap-3 pt-3 border-t border-slate-100">
              <div className="w-full md:w-72 relative">
                <input
                  type="text"
                  placeholder="예약자명, 단체명, 원본시설명 검색..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-1.5 text-xs focus:ring-2 focus:ring-sky-500/20 focus:border-sky-500 outline-none pr-8"
                />
              </div>

              <div className="w-full md:w-48">
                <select
                  value={filterFacility}
                  onChange={(e) => setFilterFacility(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 text-xs focus:ring-2 focus:ring-sky-500/20 focus:border-sky-500 outline-none cursor-pointer"
                >
                  <option value="ALL">전체 표준시설 보기</option>
                  {TARGET_FACILITIES.map((fac) => (
                    <option key={fac} value={fac}>
                      {fac}
                    </option>
                  ))}
                </select>
              </div>

              <div className="text-xs text-slate-500 font-medium md:ml-auto">
                해당 일자 <strong className="text-slate-800">{activeBookings.length}</strong>개의 예약 표시 중
              </div>
            </div>
          </div>

          {/* REAL TIME TIMETABLE PREVIEW GRID */}
          <div className="space-y-3">
            <div className="flex items-center justify-between px-1">
              <div className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                  실시간 그리드 미리보기 ({currentActiveDayObj?.name || ""} - {formatMMDD(activeTabDate)})
                </span>
              </div>
              <div className="flex items-center flex-wrap gap-x-4 gap-y-1 text-[10px] text-slate-500">
                <div className="flex items-center gap-1">
                  <span className="w-3 h-3 bg-pink-300 rounded border border-pink-400"></span>
                  <span className="font-medium text-pink-700">초과 예약분 (1일 4시간 초과)</span>
                </div>
                <div className="flex items-center gap-1">
                  <span className="w-3 h-3 bg-yellow-300 rounded border border-yellow-400"></span>
                  <span>조명사용료 (19~22시)</span>
                </div>
              </div>
            </div>

            {bookings.length === 0 ? (
              <div className="bg-white rounded-2xl border border-slate-200/80 p-12 text-center shadow-sm flex flex-col items-center justify-center gap-3">
                <FileSpreadsheet className="w-12 h-12 text-slate-300" />
                <div>
                  <div className="text-sm font-bold text-slate-700">등록된 예약 내역이 존재하지 않습니다.</div>
                  <p className="text-xs text-slate-400 max-w-sm mx-auto mt-1 leading-relaxed">
                    왼쪽의 '샘플 파일 다운로드'를 누르시거나 공유누리 엑셀 파일을 업로드해 보세요!
                  </p>
                </div>
                <button
                  onClick={handleDownloadTemplate}
                  className="mt-2 bg-slate-950 text-white hover:bg-slate-900 px-4 py-2 rounded-xl text-xs font-bold transition-all"
                >
                  샘플 엑셀 파일 받기
                </button>
              </div>
            ) : (
              renderTimeGrid()
            )}
          </div>

          {/* DUPLICATE WARNINGS & STATUS ALERTS */}
          {bookings.length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-sm space-y-3">
              <h3 className="text-xs font-bold text-slate-700 flex items-center gap-1.5 border-b border-slate-100 pb-2">
                <AlertTriangle className="w-4 h-4 text-amber-500" />
                예외 점검 및 경비실 권장 알림
              </h3>

              <div className="space-y-2">
                {/* 1일 4시간 초과 예약자 감지 */}
                {exceededBookings.length > 0 ? (
                  <div className="bg-rose-50 border border-rose-200 text-rose-950 p-4 rounded-xl text-xs space-y-2">
                    <div className="font-bold flex items-center gap-1.5 text-rose-950">
                      <AlertTriangle className="w-4.5 h-4.5 text-rose-600 fill-rose-100" />
                      1일 4시간 초과 예약 감지 (취소 권장 대상)
                    </div>
                    <p className="leading-relaxed text-[11px] text-rose-700">
                      공유누리 시설 규정에 따라 한 ID당 하루 최대 4시간까지만 이용할 수 있습니다. 4시간 한도를 누적 초과한 예약 건이 감지되었습니다. <strong>형광 분홍색</strong>으로 강조되며, 즉시 취소 처리할 수 있습니다.
                    </p>
                    <div className="space-y-1.5 pt-1 max-h-60 overflow-y-auto pr-1">
                      {exceededBookings.map((b) => (
                        <div
                          key={b.id}
                          className="bg-white/90 p-2.5 rounded border border-rose-150 flex items-center justify-between font-mono text-[10px] shadow-sm hover:border-rose-300 transition-colors"
                        >
                          <div className="space-y-0.5">
                            <div className="flex items-center gap-1">
                              <span className="bg-rose-100 text-rose-800 text-[9px] font-bold px-1.5 py-0.5 rounded">
                                {formatMMDD(b.date)} ({getDayOfWeekKorean(b.date)})
                              </span>
                              <strong className="text-slate-800 text-xs">{b.bookerName}</strong>
                              {b.groupName && <span className="text-slate-500">({b.groupName})</span>}
                            </div>
                            <div className="text-slate-600">
                              <strong>{b.resolvedFacility}</strong> | {b.startTime}~{b.endTime}
                            </div>
                            <div className="text-[9px] text-rose-600 font-sans font-medium">
                              ⚠️ 누적 4시간 초과 예약분
                            </div>
                          </div>
                          <button
                            onClick={() => handleDeleteBooking(b.id)}
                            className="bg-rose-600 hover:bg-rose-700 text-white font-sans text-[10px] font-bold px-2.5 py-1.5 rounded-lg shadow-sm flex items-center gap-1 transition-all cursor-pointer"
                            title="취소 처리"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            즉시 취소
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                ) : (
                  <div className="bg-slate-50 border border-slate-200 text-slate-600 p-3.5 rounded-xl text-xs flex items-center gap-2">
                    <CheckCircle className="w-4.5 h-4.5 text-emerald-500 flex-shrink-0" />
                    <span>모든 예약자가 1일 4시간 정책 제한 한도 내에서 올바르게 예약하였습니다.</span>
                  </div>
                )}

                {/* Night shift indicator */}
                <div className="bg-amber-50 border border-amber-200 text-amber-950 p-3.5 rounded-xl text-xs space-y-1.5">
                  <div className="font-bold flex items-center gap-1.5 text-amber-950">
                    <Zap className="w-4 h-4 text-amber-600 fill-amber-300" />
                    시즌별 조명사용료 부과 시간 자동 하이라이트 가동
                  </div>
                  <p className="leading-relaxed text-[11px] text-amber-800">
                    월별 조명사용료 규정(5~8월 저녁 19~22시 / 9~4월 저녁 18~22시 / 10~3월 새벽 06~08시 부과)에 따라 해당하는 예약 건은 엑셀 파일 및 웹 화면에서 <strong>형광 노란색(#FFFF00)</strong>으로 자동 하이라이트 처리됩니다.
                  </p>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* FOOTER */}
      <footer className="bg-white border-t border-slate-200 py-6 px-6 mt-12 text-center text-xs text-slate-400 font-mono">
        <div>공유누리 예약 관리 자동화 도구 © 2026. 100% 브라우저 자바스크립트 구동.</div>
      </footer>

      {/* MANUAL ADDITION / EDIT DIALOG */}
      <AnimatePresence>
        {isFormOpen && (
          <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-2xl border border-slate-200 shadow-xl max-w-md w-full overflow-hidden"
            >
              <div className="bg-slate-950 text-white p-4 flex items-center justify-between">
                <h3 className="font-bold font-display text-sm">
                  {editingId ? "예약 정보 상세 수정" : "신규 예약 수동 추가"}
                </h3>
                <button
                  onClick={() => setIsFormOpen(false)}
                  className="text-slate-400 hover:text-white transition-colors text-sm font-semibold px-2"
                >
                  ✕
                </button>
              </div>

              <form onSubmit={handleSaveForm} className="p-5 space-y-4 text-xs">
                {/* Facility */}
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">표준 시설 선택</label>
                  <select
                    value={formFacility}
                    onChange={(e) => setFormFacility(e.target.value as TargetFacility)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none focus:ring-2 focus:ring-sky-500/20"
                  >
                    {TARGET_FACILITIES.map((fac) => (
                      <option key={fac} value={fac}>
                        {fac}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Booker & Group */}
                <div className="grid grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">예약자명 *</label>
                    <input
                      type="text"
                      placeholder="예: 홍길동"
                      value={formBooker}
                      onChange={(e) => setFormBooker(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none focus:ring-2 focus:ring-sky-500/20"
                      required
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">단체명/소속</label>
                    <input
                      type="text"
                      placeholder="예: 클럽A"
                      value={formGroup}
                      onChange={(e) => setFormGroup(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none focus:ring-2 focus:ring-sky-500/20"
                    />
                  </div>
                </div>

                {/* Date Selection */}
                <div className="space-y-1">
                  <label className="font-bold text-slate-700">예약 일자</label>
                  <div className="flex flex-wrap gap-1.5">
                    {availableDays.map((day) => (
                      <button
                        type="button"
                        key={day.date}
                        onClick={() => setFormDate(day.date)}
                        className={`py-1.5 px-2.5 rounded-lg border text-center font-bold text-[11px] transition-all ${
                          formDate === day.date
                            ? "bg-slate-900 text-white border-slate-900 shadow-sm"
                            : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                        }`}
                      >
                        {day.name} ({formatMMDD(day.date)})
                      </button>
                    ))}
                  </div>
                </div>

                {/* Time Selection */}
                <div className="grid grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">시작 시간</label>
                    <select
                      value={formStartTime}
                      onChange={(e) => setFormStartTime(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none"
                    >
                      {hoursOptions.map((t) => (
                        <option key={t} value={t}>
                          {t}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">종료 시간</label>
                    <select
                      value={formEndTime}
                      onChange={(e) => setFormEndTime(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none"
                    >
                      {hoursOptions.concat(["22:00", "23:00"]).map((t) => (
                        <option key={t} value={t}>
                          {t}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Phone & Status */}
                <div className="grid grid-cols-2 gap-3">
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">연락처</label>
                    <input
                      type="text"
                      placeholder="010-0000-0000"
                      value={formPhone}
                      onChange={(e) => setFormPhone(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none focus:ring-2 focus:ring-sky-500/20"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="font-bold text-slate-700">예약 상태</label>
                    <select
                      value={formStatus}
                      onChange={(e) => setFormStatus(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-lg p-2 outline-none"
                    >
                      <option value="예약완료">예약완료</option>
                      <option value="승인대기">승인대기</option>
                    </select>
                  </div>
                </div>

                {/* Form Buttons */}
                <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setIsFormOpen(false)}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg font-bold transition-all"
                  >
                    취소
                  </button>
                  <button
                    type="submit"
                    className="bg-indigo-600 hover:bg-indigo-500 text-white px-5 py-2 rounded-lg font-bold transition-all shadow-md shadow-indigo-600/15"
                  >
                    저장하기
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
