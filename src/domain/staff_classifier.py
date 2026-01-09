"""
Staff Classifier Module

Handles loading and classifying staff members from CSV files.
"""

import csv
from pathlib import Path
from typing import List, Dict, Tuple

from .entities import Staff, StaffType


class StaffClassifier:
    """
    Classifies staff members based on CSV data.
    
    The CSV should have columns: Name, Type
    where Type is either "內勤" (Internal) or "外勤" (External)
    """
    
    TYPE_MAPPING = {
        "內勤": StaffType.INTERNAL,
        "internal": StaffType.INTERNAL,
        "外勤": StaffType.EXTERNAL,
        "external": StaffType.EXTERNAL,
    }
    
    def __init__(self):
        self._staff_list: List[Staff] = []
        self._internal_staff: List[Staff] = []
        self._external_staff: List[Staff] = []
    
    def load_from_csv(self, csv_path: Path) -> Tuple[List[Staff], List[Staff]]:
        """
        Load staff list from CSV file.
        
        Args:
            csv_path: Path to the CSV file
            
        Returns:
            Tuple of (internal_staff, external_staff) lists
        """
        self._staff_list = []
        self._internal_staff = []
        self._external_staff = []
        
        if not csv_path.exists():
            return ([], [])
        
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row.get('Name', row.get('name', row.get('姓名', ''))).strip()
                type_str = row.get('Type', row.get('type', row.get('類型', row.get('類別', '')))).strip()
                
                if not name:
                    continue
                
                staff_type = self.TYPE_MAPPING.get(type_str, self.TYPE_MAPPING.get(type_str.lower(), StaffType.INTERNAL))
                staff = Staff(name=name, staff_type=staff_type)
                
                self._staff_list.append(staff)
                if staff_type == StaffType.INTERNAL:
                    self._internal_staff.append(staff)
                else:
                    self._external_staff.append(staff)
        
        return (self._internal_staff, self._external_staff)
    
    def classify_from_names(
        self, 
        names: List[str], 
        known_staff: Dict[str, StaffType]
    ) -> Tuple[List[Staff], List[Staff]]:
        """
        Classify a list of names using known staff mappings.
        
        Args:
            names: List of staff names to classify
            known_staff: Dictionary mapping names to StaffType
            
        Returns:
            Tuple of (internal_staff, external_staff) lists
        """
        internal = []
        external = []
        
        for name in names:
            name = name.strip()
            if not name:
                continue
                
            staff_type = known_staff.get(name, StaffType.INTERNAL)
            staff = Staff(name=name, staff_type=staff_type)
            
            if staff_type == StaffType.INTERNAL:
                internal.append(staff)
            else:
                external.append(staff)
        
        self._internal_staff = internal
        self._external_staff = external
        self._staff_list = internal + external
        
        return (internal, external)
    
    @property
    def all_staff(self) -> List[Staff]:
        """Get all loaded staff."""
        return self._staff_list
    
    @property
    def internal_staff(self) -> List[Staff]:
        """Get internal staff only."""
        return self._internal_staff
    
    @property
    def external_staff(self) -> List[Staff]:
        """Get external staff only."""
        return self._external_staff
    
    def get_staff_by_name(self, name: str) -> Staff | None:
        """Find staff by name."""
        for staff in self._staff_list:
            if staff.name == name:
                return staff
        return None

    def add_staff(self, name: str, staff_type: StaffType, csv_path: Path) -> bool:
        """
        Append a new staff member to the CSV file.
        
        Args:
            name: Staff name
            staff_type: Staff type (Internal/External)
            csv_path: Path to the CSV file
            
        Returns:
            True if successful, False otherwise
        """
        type_str = "內勤" if staff_type == StaffType.INTERNAL else "外勤"
        
        try:
            file_exists = csv_path.exists()
            
            with open(csv_path, 'a', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                
                # If file is new or empty, write header first
                if not file_exists or csv_path.stat().st_size == 0:
                    writer.writerow(['Name', 'Type'])
                
                writer.writerow([name, type_str])
            
            # Refresh internal list
            staff = Staff(name=name, staff_type=staff_type)
            self._staff_list.append(staff)
            if staff_type == StaffType.INTERNAL:
                self._internal_staff.append(staff)
            else:
                self._external_staff.append(staff)
                
            return True
            
        except Exception as e:
            # In a real app we might log this
            print(f"Failed to append to CSV: {e}")
            return False
