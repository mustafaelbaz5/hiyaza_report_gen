from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class AnalysisResult:
    source_file: Path
    sheet_name: str
    association_name: str
    association_type: str
    village_name: str
    total_parcels: int
    basin_count: int
    registered_count: int
    unregistered_count: int
    non_held_count: int
    registered_percentage: float
    unregistered_percentage: float
    non_held_percentage: float
    detected_notes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class ReportInput:
    analysis: AnalysisResult
    association_name: str
    association_type: str
    association_code: str
    village_name: str
    delivery_date: str
    special_notes: str
    recipient: str = "Digital Smart Tree (DST)"
    executor: str = "Civil"
    report_status: str = "نهائي"

