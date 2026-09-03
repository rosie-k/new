/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface Booking {
  id: string;
  bookingNo?: string;       // 예약번호
  facilityName: string;      // 원본 시설명
  resolvedFacility: string;  // 매핑된 8개 표준 시설명 (테니스장A, 테니스장B, 테니스장C, 테니스장D, 복합운동장, 농구장A, 농구장B, 농구장C)
  bookerName: string;        // 예약자명 (혹은 단체명)
  groupName?: string;        // 단체명
  date: string;              // 예약일자 (YYYY-MM-DD)
  startTime: string;         // 시작 시간 (HH:MM)
  endTime: string;           // 종료 시간 (HH:MM)
  phoneNumber?: string;      // 연락처
  status?: string;           // 상태 (예약완료, 승인대기 등)
  isNightHighlighted?: boolean; // 야간 하이라이트 여부 (19시~22시 포함 여부)
  isDuplicate?: boolean;     // 중복 예약 여부
  isExceeded?: boolean;      // 1일 4시간 초과 여부
  exceededNote?: string;     // 초과 사유 노트 ("초과예약분")
}

export const TARGET_FACILITIES = [
  "테니스장A",
  "테니스장B",
  "테니스장C",
  "테니스장D",
  "복합운동장",
  "농구장A",
  "농구장B",
  "농구장C"
] as const;

export type TargetFacility = typeof TARGET_FACILITIES[number];

export interface AutomationState {
  status: "IDLE" | "CONNECTING" | "WAITING_LOGIN" | "OTP_VERIFIED" | "NAVIGATING" | "DOWNLOADING" | "COMPLETED" | "ERROR";
  message: string;
  progress: number;
}
