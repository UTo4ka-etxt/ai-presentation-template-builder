"""Native PPTX assembly from a decomposed template package."""

from .catalog import TemplateCatalog
from .decompose import PptxDecomposer
from .ooxml import PresentationBuilder

__all__ = ["PresentationBuilder", "PptxDecomposer", "TemplateCatalog"]
