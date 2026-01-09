"""
Sorting Utilities Module

Provides sorting functions for attendance data output.
"""

from typing import List
from domain.entities import MonthlyAttendance


def count_strokes(char: str) -> int:
    """
    Get stroke count for a Chinese character.
    Uses Unicode CJK Strokes data approximation.
    
    This is a simplified implementation that groups by common stroke ranges.
    For production use, consider using a proper stroke database.
    """
    # Common Chinese surname first characters with stroke counts
    STROKE_MAP = {
        # 1-4 strokes
        '丁': 2, '王': 4, '毛': 4, '方': 4, '文': 4, '孔': 4, '牛': 4,
        # 5-6 strokes
        '白': 5, '石': 5, '田': 5, '史': 5, '左': 5, '古': 5, '司': 5, '甘': 5,
        '朱': 6, '江': 6, '向': 6, '任': 6, '伍': 6, '池': 6, '安': 6,
        # 7 strokes
        '李': 7, '吳': 7, '何': 7, '余': 7, '宋': 7, '呂': 7, '杜': 7, '沈': 7, '汪': 7,
        '巫': 7, '辛': 7, '阮': 7, '邱': 7, '吕': 7, '冷': 7, '沙': 7,
        # 8 strokes
        '林': 8, '周': 8, '金': 8, '邵': 8, '武': 8, '范': 8, '卓': 8, '侯': 8,
        '易': 8, '尚': 8, '洪': 8, '祁': 8, '柯': 8, '柏': 8, '施': 8,
        # 9 strokes
        '張': 11, '陳': 11, '柳': 9, '洪': 9, '胡': 9, '姚': 9, '紀': 9, '俞': 9,
        '段': 9, '祝': 9, '侯': 9, '姜': 9, '柳': 9, '封': 9, '查': 9,
        # 10 strokes
        '孫': 10, '高': 10, '徐': 10, '馬': 10, '郭': 10, '唐': 10, '倪': 10, '凌': 10,
        '翁': 10, '夏': 10, '殷': 10, '秦': 10, '袁': 10, '涂': 10, '翁': 10,
        # 11-12 strokes
        '張': 11, '陳': 11, '許': 11, '曹': 11, '梁': 11, '莊': 11, '康': 11, '郭': 11,
        '黃': 12, '曾': 12, '程': 12, '彭': 12, '傅': 12, '富': 12, '游': 12,
        # 13-15 strokes
        '楊': 13, '葉': 13, '董': 13, '溫': 13, '詹': 13, '雷': 13, '廖': 14,
        '趙': 14, '熊': 14, '劉': 15, '蔣': 15, '蔡': 15, '鄭': 15, '蕭': 15,
        # 16+ strokes
        '賴': 16, '錢': 16, '盧': 16, '謝': 17, '鍾': 17, '戴': 17, '鄧': 17,
        '蘇': 19, '羅': 19, '魏': 18, '龔': 22, '龍': 16,
    }
    
    if char in STROKE_MAP:
        return STROKE_MAP[char]
    
    # Fallback: estimate by Unicode code point range
    # This gives a rough ordering for characters not in the map
    code = ord(char)
    if 0x4E00 <= code <= 0x9FFF:  # CJK Unified Ideographs
        # Simple heuristic based on code point (not accurate but provides ordering)
        return ((code - 0x4E00) % 20) + 5
    
    return 10  # Default for non-CJK characters


def get_name_stroke_key(name: str) -> tuple:
    """
    Get sort key for name based on first character stroke count.
    Returns tuple of (stroke_count, full_name) for stable sorting.
    """
    if not name:
        return (0, "")
    first_char = name[0]
    return (count_strokes(first_char), name)


def sort_attendance_list(
    attendance_list: List[MonthlyAttendance],
    sort_by: str = "attendance_rate"
) -> List[MonthlyAttendance]:
    """
    Sort attendance list by specified criteria.
    
    Args:
        attendance_list: List of MonthlyAttendance objects
        sort_by: Sorting method - "attendance_rate" or "name_strokes"
        
    Returns:
        Sorted list (new list, does not modify original)
    """
    if sort_by == "name_strokes":
        # Sort by surname stroke count (ascending), then by full name
        return sorted(
            attendance_list,
            key=lambda m: get_name_stroke_key(m.staff.name)
        )
    else:
        # Default: sort by attendance rate (descending - higher rate first)
        return sorted(
            attendance_list,
            key=lambda m: -m.attendance_rate
        )
