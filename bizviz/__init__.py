"""BizViz — turn an Excel workbook into an openable Power BI project."""

from .build import BuildResult, build_project
from .design import ReportSpec, claude_design, rule_based_design
from .modeling import SemanticModel, build_model
from .profiling import TableProfile, profile_workbook

__all__ = [
    "profile_workbook",
    "TableProfile",
    "build_model",
    "SemanticModel",
    "rule_based_design",
    "claude_design",
    "ReportSpec",
    "build_project",
    "BuildResult",
]

__version__ = "0.1.0"
