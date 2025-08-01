from flask import Flask, render_template, request, jsonify, send_file
import os
import json
import ast
import random
import re
import traceback
from datetime import datetime
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
from openai import OpenAI
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (
    XL_CHART_TYPE,
    XL_LEGEND_POSITION,
    XL_TICK_LABEL_POSITION,
    XL_TICK_MARK,
)
from enum import Enum
from typing import Optional, Dict, List, Any

# Initialize Flask app
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = 'generated_ppts'
app.config['TEMPLATE_FOLDER'] = 'templates'

# Create necessary directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMPLATE_FOLDER'], exist_ok=True)

# Load environment variables
load_dotenv()
api_key_openAI = os.getenv("OPENAI_API_KEY")
api_key_gemini = os.getenv("GEMINI_API_KEY")

# Initialize API clients
client = OpenAI(api_key=api_key_openAI)
genai.configure(api_key=api_key_gemini)

# --- AI REQUEST TYPES ---
class AIRequestType(Enum):
    COMPANY_LIST = "company_list"
    EXECUTIVE_SUMMARY = "executive_summary"
    MARKET_ENABLERS = "market_enablers"
    INDUSTRY_EXPANSION = "industry_expansion"
    INVESTMENT_CHALLENGES = "investment_challenges"
    COMPANY_INFO = "company_info"
    COMPANY_REVENUE = "company_revenue"

# --- FALLBACK RESPONSES ---
FALLBACK_RESPONSES = {
    AIRequestType.COMPANY_LIST: [
        "TechCorp Inc.", "Global Systems Ltd.", "Innovation Partners", 
        "Market Leaders Co.", "Industry Solutions", "Premier Tech",
        "Advanced Systems", "Future Innovations", "Digital Dynamics", "Tech Pioneers"
    ],
    AIRequestType.EXECUTIVE_SUMMARY: "The {headline} market is experiencing robust growth, with a projected CAGR of 12.5% from {base_year} to {forecast_year}. The market size, valued at {cur} {rev_current} {value_in} in {base_year}, is expected to reach {cur} {rev_future} {value_in} by {forecast_year}. Key growth drivers include technological advancement, increasing demand, and market expansion in emerging economies.",
    AIRequestType.MARKET_ENABLERS: "Digital Transformation: The rapid digitalization across industries is driving unprecedented demand for {headline}, with organizations investing heavily in modernization initiatives to remain competitive.\n\nRegulatory Support: Government initiatives and favorable policies are creating a conducive environment for market growth, with increased funding and tax incentives supporting industry expansion.",
    AIRequestType.INDUSTRY_EXPANSION: "The {main_topic} segment is witnessing unprecedented growth, driven by technological innovations and increasing adoption across various industries. This expansion is creating significant opportunities for market players to develop specialized solutions.\n\nMarket demand for {main_topic} solutions continues to surge as organizations recognize the strategic value and operational benefits. The increasing investment in research and development is accelerating product innovation and market penetration.",
    AIRequestType.INVESTMENT_CHALLENGES: "The substantial capital requirements for implementing {headline} solutions present a significant barrier for small and medium enterprises. High upfront costs for infrastructure, technology, and skilled personnel limit market entry.\n\nReturn on investment concerns continue to challenge market adoption, as organizations struggle to justify the initial expenditure against uncertain long-term benefits. This financial hesitation particularly impacts emerging markets where capital availability is constrained.",
    AIRequestType.COMPANY_INFO: {
        "company_name": "{company_name}",
        "headquarters": "New York, USA",
        "employee_count": "10,000+",
        "top_product": "Advanced Technology Solutions",
        "estd": "1995",
        "website": "www.example.com",
        "geographic_presence": "Global - 50+ countries",
        "ownership": "Public (NYSE: TECH)",
        "short_description_company": "{company_name} is a leading global provider in the {headline} market, offering innovative solutions that drive digital transformation across industries. With over 25 years of experience, the company has established itself as a trusted partner for organizations seeking to leverage cutting-edge technology. The company's comprehensive portfolio includes advanced analytics, cloud solutions, and enterprise software that enable businesses to optimize operations and achieve sustainable growth. Through continuous innovation and strategic partnerships, {company_name} maintains its position at the forefront of technological advancement, serving thousands of clients worldwide."
    },
    AIRequestType.COMPANY_REVENUE: [
        random.randint(80, 120),
        random.randint(90, 130),
        random.randint(100, 150)
    ]
}

# --- UNIFIED AI FUNCTION ---
class AIService:
    def __init__(self, openai_client, gemini_api_key):
        self.openai_client = openai_client
        self.gemini_configured = False
        if gemini_api_key:
            genai.configure(api_key=gemini_api_key)
            self.gemini_configured = True
    
    def generate_content(self, request_type: AIRequestType, context: Dict[str, Any]) -> Any:
        """
        Unified AI content generation function with fallback handling
        
        Args:
            request_type: Type of AI request from AIRequestType enum
            context: Dictionary containing context variables for the request
            
        Returns:
            Generated content or fallback response
        """
        try:
            if request_type == AIRequestType.COMPANY_LIST:
                return self._generate_company_list(context)
            elif request_type == AIRequestType.EXECUTIVE_SUMMARY:
                return self._generate_executive_summary(context)
            elif request_type == AIRequestType.MARKET_ENABLERS:
                return self._generate_market_enablers(context)
            elif request_type == AIRequestType.INDUSTRY_EXPANSION:
                return self._generate_industry_expansion(context)
            elif request_type == AIRequestType.INVESTMENT_CHALLENGES:
                return self._generate_investment_challenges(context)
            elif request_type == AIRequestType.COMPANY_INFO:
                return self._generate_company_info(context)
            elif request_type == AIRequestType.COMPANY_REVENUE:
                return self._generate_company_revenue(context)
        except Exception as e:
            print(f"AI generation failed for {request_type.value}: {str(e)}")
            return self._get_fallback_response(request_type, context)
    
    def _generate_company_list(self, context: Dict[str, Any]) -> List[str]:
        prompt = f'Based on the headline "{context["headline"]}", generate a list of 10 relevant company names. Only return a Python list in the format: ["Company 1", "Company 2"]'
        response = self.openai_client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return ast.literal_eval(response.choices[0].message.content.strip())
    
    def _generate_executive_summary(self, context: Dict[str, Any]) -> str:
        prompt = f"Write an executive summary for {context['headline']} using components like CAGR and market share with expected growth rate and revenue, for a global market of {context['value_in']} in {context['cur']} for the years {context['historical_year']} to {context['forecast_year']} within 125 words. Return only a single paragraph."
        response = self.openai_client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    
    def _generate_market_enablers(self, context: Dict[str, Any]) -> str:
        prompt = f'Write an executive summary about key market enablers (2 points) for {context["headline"]}, each 100-125 words. Return a Python list like ["heading: context", "heading: context"].'
        response = self.openai_client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return "\n".join(ast.literal_eval(response.choices[0].message.content))
    
    def _generate_industry_expansion(self, context: Dict[str, Any]) -> str:
        prompt = f'Write about "The Expanding {context["main_topic"]} Industry As A Key Driver For {context["headline"]}" (2 points, 90-120 words each). Return a Python list like ["content", "content"].'
        response = self.openai_client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return "\n".join(ast.literal_eval(response.choices[0].message.content))
    
    def _generate_investment_challenges(self, context: Dict[str, Any]) -> str:
        prompt = f'Write about "High Initial Investment And Setup Costs Are Posing Challenge To The Growth Of The Market for {context["headline"]}" (2 points, 90-120 words each). Return a Python list like ["content", "content"].'
        response = self.openai_client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return "\n".join(ast.literal_eval(response.choices[0].message.content))
    
    def _generate_company_info(self, context: Dict[str, Any]) -> Dict[str, str]:
        prompt = f'{context["company_name"]} is a real company in the "{context["headline"]}" domain. Return info in JSON format: {{"company_name": "", "headquarters": "", "employee_count": "", "top_product": "", "estd": "", "website": "", "geographic_presence": "", "ownership": "", "short_description_company": "(around 200 words)"}}. Only return valid JSON.'
        response = self.openai_client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return json.loads(response.choices[0].message.content)
    
    def _generate_company_revenue(self, context: Dict[str, Any]) -> List[int]:
        if not self.gemini_configured:
            return self._get_fallback_response(AIRequestType.COMPANY_REVENUE, context)
        
        model = genai.GenerativeModel("gemini-2.0-flash")
        prompt = f"Give me only a list of estimated revenue in {context['currency']} for {context['company_name']} for 2022, 2023, and 2024. Format: [value_2022, value_2023, value_2024]. If no data is available, generate realistic random numbers. No extra text no explanation just the list."
        response = model.generate_content(prompt)
        return ast.literal_eval(response.text.strip())
    
    def _get_fallback_response(self, request_type: AIRequestType, context: Dict[str, Any]) -> Any:
        """Get fallback response with context variable substitution"""
        fallback = FALLBACK_RESPONSES[request_type]
        
        if isinstance(fallback, str):
            # Replace placeholders in string
            for key, value in context.items():
                fallback = fallback.replace(f"{{{key}}}", str(value))
            return fallback
        elif isinstance(fallback, dict):
            # Deep copy and replace placeholders in dictionary
            result = {}
            for k, v in fallback.items():
                if isinstance(v, str):
                    for key, value in context.items():
                        v = v.replace(f"{{{key}}}", str(value))
                result[k] = v
            return result
        else:
            # Return as-is for lists and other types
            return fallback

# Initialize AI Service globally
ai_service = AIService(client, api_key_gemini)

# --- PRESENTATION MODIFICATION CLASSES & FUNCTIONS ---
class TaxonomyBoxGenerator:
    COLORS = {
        "purple": RGBColor(0x31, 0x09, 0x7E),
        "orange": RGBColor(255, 102, 51),
        "teal": RGBColor(0, 179, 152),
        "blue": RGBColor(0, 162, 232),
        "dark_blue": RGBColor(36, 64, 142),
        "white": RGBColor(255, 255, 255),
        "light_gray": RGBColor(0xF2, 0xF2, 0xF2),
        "text_dark": RGBColor(0, 0, 0),
    }
    CATEGORY_COLORS = {"DEFAULT": COLORS["purple"], "BY REGION": COLORS["dark_blue"]}

    def __init__(self, presentation):
        self.prs = presentation
        self.slide_width = self.prs.slide_width
        self.slide_height = self.prs.slide_height
        self.left_margin, self.top_margin, self.right_margin, self.bottom_margin = (
            Inches(0.5),
            Inches(2),
            Inches(0.5),
            Inches(0.8),
        )
        self.h_spacing, self.v_spacing = Inches(0.2), Inches(0.2)

    def _add_category_box(
        self, slide, category, content, left, top, max_width, max_height
    ):
        header_color = self.CATEGORY_COLORS.get(category, self.COLORS["purple"])
        header_height = Inches(0.3)
        header = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, max_width, header_height
        )
        header.fill.solid()
        header.fill.fore_color.rgb = header_color
        header.line.color.rgb = header_color
        p = header.text_frame.paragraphs[0]
        p.text = category
        p.font.size, p.font.bold, p.font.color.rgb, p.alignment = (
            Pt(11),
            True,
            self.COLORS["white"],
            PP_ALIGN.CENTER,
        )

        content_box_height = max_height - header_height + Inches(0.2)
        content_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left,
            top + header_height,
            max_width,
            content_box_height,
        )
        content_box.fill.solid()
        content_box.fill.fore_color.rgb = self.COLORS["light_gray"]
        content_box.line.color.rgb = self.COLORS["light_gray"]
        tf = content_box.text_frame
        tf.word_wrap, tf.vertical_anchor = True, MSO_VERTICAL_ANCHOR.TOP
        self._add_list_content(tf, content)

    def _add_list_content(self, text_frame, content):
        """Add bullet list content with hollow bullets"""
        text_frame.margin_bottom = Pt(12)
        if text_frame.paragraphs:
            text_frame.paragraphs[0].text = ""

        p = text_frame.paragraphs[0]
        pPr = p._p.get_or_add_pPr()

        lst = pPr.find(qn("a:lstStyle"))
        if lst is None:
            lst = OxmlElement("a:lstStyle")
            pPr.append(lst)

        for i, item in enumerate(content):
            p = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
            p.text = item
            p.alignment = PP_ALIGN.LEFT

            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", str(int(Pt(15).emu)))
            pPr.set("indent", str(int(Pt(-15).emu)))

            marL = OxmlElement("a:marL")
            marL.set("val", str(int(Pt(50).emu)))
            pPr.append(marL)

            indent = OxmlElement("a:indent")
            indent.set("val", str(int(Pt(-19).emu)))
            pPr.append(indent)

            buChar = OxmlElement("a:buChar")
            buChar.set("char", "○")
            pPr.append(buChar)

            buFont = OxmlElement("a:buFont")
            buFont.set("typeface", "Symbol")
            pPr.append(buFont)

            buClr = OxmlElement("a:buClr")
            srgbClr = OxmlElement("a:srgbClr")
            srgbClr.set("val", "000000")
            buClr.append(srgbClr)
            pPr.append(buClr)

            for run in p.runs:
                run.font.size = Pt(10)
                run.font.color.rgb = self.COLORS["text_dark"]

            if not pPr.find(qn("a:spcAft")):
                spcAft = OxmlElement("a:spcAft")
                spcVal = OxmlElement("a:spcPts")
                spcVal.set("val", "600")
                spcAft.append(spcVal)
                pPr.append(spcAft)

            if not pPr.find(qn("a:lnSpc")):
                lnSpc = OxmlElement("a:lnSpc")
                spcPts = OxmlElement("a:spcPts")
                spcPts.set("val", "1800")
                lnSpc.append(spcPts)
                pPr.append(lnSpc)

    def add_taxonomy_boxes(self, slide_index, taxonomy_data):
        slide = self.prs.slides[slide_index]
        available_width = self.slide_width - self.left_margin - self.right_margin
        num_categories = len(taxonomy_data)
        boxes_per_row = min(5, num_categories)
        box_width = (
            available_width - (boxes_per_row - 1) * self.h_spacing
        ) / boxes_per_row

        rows, current_row, current_row_width = [], [], 0
        for category, hierarchy in taxonomy_data.items():
            item_count = len(hierarchy)
            box_height = max(
                Inches(1), Inches(0.43) + (item_count * Inches(0.17) * 1.2)
            )
            if current_row_width + box_width > available_width and current_row:
                rows.append(current_row)
                current_row, current_row_width = [], 0
            current_row.append(
                {"category": category, "content": hierarchy, "height": box_height}
            )
            current_row_width += box_width + self.h_spacing
        if current_row:
            rows.append(current_row)

        current_top = self.top_margin
        for row in rows:
            row_max_height = max(box["height"] for box in row)
            row_width = len(row) * box_width + (len(row) - 1) * self.h_spacing
            left_start = self.left_margin + (available_width - row_width) / 2
            for i, box in enumerate(row):
                left = left_start + i * (box_width + self.h_spacing)
                self._add_category_box(
                    slide,
                    box["category"],
                    box["content"],
                    left,
                    current_top,
                    box_width,
                    box["height"],
                )
            current_top += row_max_height + self.v_spacing


def replace_text_in_presentation(prs, slide_data_dict):
    """Replaces placeholders in the entire presentation."""
    for slide_idx, slide in enumerate(prs.slides):
        data = slide_data_dict.get(slide_idx, {})
        if not data:
            continue
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                for key, value in data.items():
                    token = f"{{{{{key}}}}}"
                    if token in p.text:
                        inline_text = p.text
                        p.text = inline_text.replace(token, str(value))


def replace_text_in_tables(prs, slide_indices, slide_data_dict):
    """Replaces placeholders specifically within tables on given slides."""
    for idx in slide_indices:
        if idx >= len(prs.slides):
            continue
        slide = prs.slides[idx]
        data = slide_data_dict.get(idx, {})
        if not data:
            continue
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            for row in shape.table.rows:
                for cell in row.cells:
                    for p in cell.text_frame.paragraphs:
                        for run in p.runs:
                            for key, value in data.items():
                                if f"{{{{{key}}}}}" in run.text:
                                    run.text = run.text.replace(
                                        f"{{{{{key}}}}}", str(value)
                                    )


def get_rgb_color_safe(font):
    """Safely get the RGB color from a font, returns None if not explicitly set."""
    try:
        return font.color.rgb
    except AttributeError:
        return None


def replace_text_preserving_color(paragraph, placeholder, new_text):
    full_text = "".join(run.text for run in paragraph.runs)

    if placeholder not in full_text:
        return

    for run in paragraph.runs:
        if placeholder in run.text:
            font_color = get_rgb_color_safe(run.font)
            run.text = run.text.replace(placeholder, new_text)
            if font_color:
                run.font.color.rgb = font_color
            break


def replace_text_in_paragraph(paragraph, placeholder, new_text):
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder in full_text:
        new_full_text = full_text.replace(placeholder, new_text)
        for run in paragraph.runs:
            run.text = ""
        paragraph.runs[0].text = new_full_text


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for line_dir in ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]:
        ln = OxmlElement(line_dir)
        ln.set("w", "12700")

        solidFill = OxmlElement("a:solidFill")
        srgbClr = OxmlElement("a:srgbClr")
        srgbClr.set("val", "A6A6A6")
        solidFill.append(srgbClr)
        ln.append(solidFill)

        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")

        tcPr.append(ln)


def validate_segment_hierarchy(segment_text):
    """Validate the segment hierarchy structure"""
    lines = segment_text.strip().split('\n')
    errors = []
    last_main_number = 0
    last_sub_numbers = {}

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        match = re.match(r'^(\d+(?:\.\d+)*)\.\s+(.+)$', line)
        if not match:
            errors.append(f"Line {i + 1}: Invalid format")
            continue

        number_parts = [int(n) for n in match.group(1).split('.')]
        depth = len(number_parts)

        if depth == 1:
            if number_parts[0] != last_main_number + 1:
                errors.append(f"Line {i + 1}: Expected main number {last_main_number + 1}, got {number_parts[0]}")
            last_main_number = number_parts[0]
            last_sub_numbers = {}
        elif depth == 2:
            main_num = number_parts[0]
            sub_num = number_parts[1]
            
            if main_num != last_main_number:
                errors.append(f"Line {i + 1}: Sub-item doesn't match current main number")
            
            expected_sub = last_sub_numbers.get(main_num, 0) + 1
            if sub_num != expected_sub:
                errors.append(f"Line {i + 1}: Expected sub-number {main_num}.{expected_sub}")
            last_sub_numbers[main_num] = sub_num
        elif depth == 3:
            main_num = number_parts[0]
            sub_num = number_parts[1]
            sub_sub_num = number_parts[2]
            
            if main_num != last_main_number:
                errors.append(f"Line {i + 1}: Sub-sub-item doesn't match current main number")
            
            key = f"{main_num}.{sub_num}"
            expected_sub_sub = last_sub_numbers.get(key, 0) + 1
            if sub_sub_num != expected_sub_sub:
                errors.append(f"Line {i + 1}: Expected sub-sub-number {main_num}.{sub_num}.{expected_sub_sub}")
            last_sub_numbers[key] = sub_sub_num

    return errors


def generate_random_data():
    """Generate random data for charts"""
    years = list(range(2019, 2033))
    type_ranges = {1: (40, 50), 2: (25, 35), 3: (15, 25), 4: (5, 25)}
    data = []
    for year in years:
        row = [year]
        for i in range(4): 
            low, high = type_ranges.get(i + 1, (5, 25))
            row.append(random.randint(low, high))
        data.append(row)
    return data


def parse_segment_input(segment_input: str) -> Dict[str, Dict]:
    """Parse segment input into a nested dictionary"""
    lines = segment_input.strip().split("\n")
    nested_dict = {}
    level_stack = []
    for line in lines:
        if not line.strip():
            continue
        key, value = line.split(". ", 1)
        parts = key.split(".")
        depth = len(parts)
        label = value.strip()
        level_stack = level_stack[: depth - 1]
        current = nested_dict
        for k in level_stack:
            current = current[k]
        current[label] = {}
        level_stack.append(label)
    return nested_dict


def generate_toc_data(nested_dict: Dict, headline: str, forecast_period: str, user_segment) -> Dict[str, int]:
    """Generate Table of Contents data"""
    toc_start_levels = {
        "1. Introduction": 0,
        "1.1. Objectives of the Study": 1,
        "1.2. Market Definition & Scope": 1,
        "2. Research Methodology": 0,
        "2.1. Research Process": 1,
        "2.2. Secondary & Primary Data Methods": 1,
        "2.3. Market Size Estimation Methods": 1,
        "2.4. Market Assumptions & Limitations": 1,
        "3. Executive Summary": 0,
        "3.1. Global Market Outlook": 1,
        "3.2. Key Market Highlights": 1,
        "3.3. Segmental Overview": 1,
        "3.4. Competition Overview": 1,
        "4. Market Dynamics & Outlook": 0,
        "4.1. Macro-Economic Indicators​": 1,
        "4.2. Drivers & Opportunitiess": 1,
        "4.3. Restraints & Challenges": 1,
        "4.4. Supply Side Trends": 1,
        "4.5. Demand Side Trends": 1,
        "4.6. Porter's Analysis & Impact": 1,
        "4.6.1. Competitive Rivalry": 2,
        "4.6.2. Threat of substitutes": 2,
        "4.6.3. Bargaining power of buyers": 2,
        "4.6.4. Threat of new entrants": 2,
        "4.6.5. Bargaining power of suppliers": 2,
        "5. Key Market Insights": 0,
        "5.1. Key Success Factors": 1,
        "5.2. Market Impacting Factors": 1,
        "5.3. Top Investment Pockets": 1,
        "5.4. Market Attractiveness Index": 1,
        "5.5. Market Ecosystem ": 1,
        "5.6. PESTEL Analysis": 1,
        "5.7. Pricing Analysis": 1,
        "5.8. Regulatory Landscape": 1,
        "5.9. Case Study Analysis": 1,
        "5.10. Supply Chain Analysis": 1,
        "5.11. Raw Material Analysis": 1,
        "5.12. Customer Buying Behaviour Analysis": 1,
    }

    toc_mid = {}
    main_index = 6
    for type_index, (type_name, points) in enumerate(nested_dict.items(), start=main_index):
        toc_mid[
            f"{type_index}. {headline} Size by {type_name} ({forecast_period})"
        ] = 0
        point_count = 1
        for point, subpoints in points.items():
            toc_mid[f"{type_index}.{point_count}. {point}"] = 1
            if subpoints:
                for sp_count, sub in enumerate(subpoints, start=1):
                    toc_mid[f"{type_index}.{point_count}.{sp_count}. {sub}"] = 2
            point_count += 1
        


    x = len(list(nested_dict.keys())) + 6
    toc_end_levels = {
        f"{x}. {headline} Size by Region ({forecast_period})": 0,
        f"{x}.1. North America ({user_segment})": 1,
        f"{x}.1.1. US": 2,
        f"{x}.1.2. Canada": 2,
        f"{x}.2. Europe ({user_segment})": 1,
        f"{x}.2.1. UK": 2,
        f"{x}.2.2. Germany": 2,
        f"{x}.2.3. Spain": 2,
        f"{x}.2.4. France": 2,
        f"{x}.2.5. Italy": 2,
        f"{x}.2.6. Rest of Europe": 2,
        f"{x}.3. Asia-Pacific ({user_segment})": 1,
        f"{x}.3.1. China": 2,
        f"{x}.3.2. India": 2,
        f"{x}.3.3. Japan": 2,
        f"{x}.3.4. South Korea": 2,
        f"{x}.3.5. Rest of Asia Pacific": 2,
        f"{x}.4. Latin America ({user_segment})": 1,
        f"{x}.4.1. Brazil": 2,
        f"{x}.4.2. Mexico": 2,
        f"{x}.4.3. Rest of Latin America": 2,
        f"{x}.5. Middle East & Africa ({user_segment})": 1,
        f"{x}.5.1. GCC Countries": 2,
        f"{x}.5.2. South Africa": 2,
        f"{x}.5.3. Rest of Middle East & Africa": 2,
        f"{x+1}. Competitive Dashboard": 0,
        f"{x+1}.1. Top 5 Player Comparison": 1,
        f"{x+1}.2. Market Positioning of Key Players, 2024": 1,
        f"{x+1}.3. Strategies Adopted by Key Market Players": 1,
        f"{x+1}.4. Recent Developments in the Market": 1,
        f"{x+1}.5. Company Market Share Analysis, 2024": 1,
        f"{x+2}. Key Company Profiles": 0,
    }
    
    return {**toc_start_levels, **toc_mid, **toc_end_levels}


def add_toc_to_slides(prs: Presentation, toc_data_levels: Dict[str, int], toc_slide_indices: List[int]):
    """Add Table of Contents to specified slides"""
    for i in toc_slide_indices:
        slide = prs.slides[i]
        table_shape = slide.shapes.add_table(
            20, 2, Inches(3.1), Inches(0.3), Inches(10), Inches(6.8)
        )
        table = table_shape.table
        for row in table.rows:
            for cell in row.cells:
                cell.text = ""
                cell.fill.background()
                tcPr = cell._tc.get_or_add_tcPr()
                for border_tag in ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]:
                    tcPr.append(
                        parse_xml(
                            f'<{border_tag} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></{border_tag}>'
                        )
                    )

    content_items = list(toc_data_levels.keys())
    content_index = 0
    for i in toc_slide_indices:
        table = prs.slides[i].shapes[-1].table
        for col in range(2):
            for row in range(20):
                if content_index >= len(content_items):
                    break
                cell, key = table.cell(row, col), content_items[content_index]
                level = toc_data_levels[key]
                para = cell.text_frame.paragraphs[0]
                para.text = "          " * level + key
                font = para.font
                font.color.rgb, font.size, font.name = RGBColor(0, 0, 0), Pt(11), "Calibri"
                if level == 0:
                    font.color.rgb, font.bold = RGBColor(112, 48, 160), True
                else:
                    font.color.rgb = RGBColor(0, 0, 0)
                    font.bold = False
                content_index += 1


def create_chart_on_slide(slide: Any, data: List[List], chart_columns: List[str], 
                         left: float, top: float, width: float, height: float):
    """Create a stacked column chart on the specified slide"""
    chart_data = CategoryChartData()
    chart_data.categories = [str(row[0]) for row in data]

    # Add series in correct order
    for i, col_name in enumerate(chart_columns[:4]):  # Limit to 4 columns
        chart_data.add_series(col_name, [row[i + 1] for row in data])

    # Add Chart
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED, left, top, width, height, chart_data
    ).chart
    chart.plots[0].gap_width = 250
    
    # Style
    chart.chart_style = 2
    chart.has_title = False

    # Legend Position
    chart.has_legend = True
    chart.legend.font.size = Pt(10)
    chart.legend.font.name = "Poppins"
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    # Remove gridlines & labels
    value_axis = chart.value_axis
    value_axis.visible = False
    value_axis.has_major_gridlines = False
    value_axis.has_minor_gridlines = False
    value_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    value_axis.major_tick_mark = XL_TICK_MARK.NONE
    value_axis.minor_tick_mark = XL_TICK_MARK.NONE

    category_axis = chart.category_axis
    category_axis.major_tick_mark = XL_TICK_MARK.NONE
    category_axis.minor_tick_mark = XL_TICK_MARK.NONE
    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.size = Pt(10)
    cat_axis.tick_labels.font.name = "Poppins"
    cat_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        form_data = request.form
        required_fields = [
            'headline', 'headline_2', 'historical_year', 'base_year',
            'forecast_year', 'forecast_period', 'cur', 'value_in',
            'rev_current', 'rev_future', 'org_1', 'org_2', 'org_3',
            'org_4', 'org_5', 'paper_1', 'paper_2', 'paper_3',
            'paper_4', 'paper_5', 'segment_input'
        ]
        
        missing_fields = []
        for field in required_fields:
            if not form_data.get(field, '').strip():
                missing_fields.append(field)
        
        if missing_fields:
            return jsonify({
                'error': 'Missing required fields',
                'fields': missing_fields
            }), 400
        
        segment_errors = validate_segment_hierarchy(form_data['segment_input'])
        if segment_errors:
            return jsonify({
                'error': 'Invalid segment hierarchy',
                'details': segment_errors
            }), 400
        
        # Extract form data
        headline = form_data['headline']
        headline_2 = headline
        historical_year = "2019-2023"
        base_year = "2024"
        forecast_year = "2032"
        forecast_period = "2025-2032"
        cur = "USD"
        value_in = form_data['value_in']
        currency = f"{cur} {value_in}"
        
        org_1 = form_data['org_1']
        org_2 = form_data['org_2']
        org_3 = form_data['org_3']
        org_4 = form_data['org_4']
        org_5 = form_data['org_5']
        
        paper_1 = form_data['paper_1']
        paper_2 = form_data['paper_2']
        paper_3 = form_data['paper_3']
        paper_4 = form_data['paper_4']
        paper_5 = form_data['paper_5']
        
        rev_current = form_data['rev_current']
        rev_future = form_data['rev_future']
        
        segment_input = form_data['segment_input']
        
        # --- Data Processing and Content Generation ---
        
        # Parse segment input
        nested_dict = parse_segment_input(segment_input)
        main_topic = list(nested_dict.keys())
        s_segment = "By " + ",\nBy ".join(main_topic)
        user_segment = "By " + ", By ".join(main_topic)

        # Generate context
        output_lines = []
        for main_type, points in nested_dict.items():
            line_parts = []
            for point, subpoints in points.items():
                if subpoints:
                    subpoint_str = ", ".join(subpoints.keys())
                    line_parts.append(f"{point} ({subpoint_str})")
                else:
                    line_parts.append(point)
            output_lines.append(f"{main_type}: {', '.join(line_parts)}")
        output_lines.append(
            "By Region: North America, Europe, Asia-Pacific, Latin America, Middle East & Africa"
        )
        context = "\n".join(output_lines)

        # Generate TOC
        toc_data_levels = generate_toc_data(nested_dict, headline, forecast_period,user_segment)
        
        # Generate AI content using unified service
        ai_context = {
            'headline': headline,
            'value_in': value_in,
            'cur': cur,
            'historical_year': historical_year,
            'forecast_year': forecast_year,
            'base_year': base_year,
            'rev_current': rev_current,
            'rev_future': rev_future,
            'main_topic': main_topic[0] if main_topic else "Type 1",
            'currency': currency
        }
        
        # Generate all AI content
        company_list = ai_service.generate_content(AIRequestType.COMPANY_LIST, ai_context)
        mpara_11 = ai_service.generate_content(AIRequestType.EXECUTIVE_SUMMARY, ai_context)
        para_11 = ai_service.generate_content(AIRequestType.MARKET_ENABLERS, ai_context)
        para_14 = ai_service.generate_content(AIRequestType.INDUSTRY_EXPANSION, ai_context)
        para_15 = ai_service.generate_content(AIRequestType.INVESTMENT_CHALLENGES, ai_context)
        
        # Get company info
        ai_context['company_name'] = company_list[0]
        company_info = ai_service.generate_content(AIRequestType.COMPANY_INFO, ai_context)
        revenue_list = ai_service.generate_content(AIRequestType.COMPANY_REVENUE, ai_context)
        
        # Add company profiles to TOC
        x = len(main_topic) + 6

        first_company = company_list[0]
        toc_data_levels[f"{x+2}.1. {first_company}"] = 1
        toc_data_levels[f"{x+2}.1.1. Company Overview"] = 2
        toc_data_levels[f"{x+2}.1.2. Product Portfolio Overview"] = 2
        toc_data_levels[f"{x+2}.1.3. Financial Overview"] = 2
        toc_data_levels[f"{x+2}.1.4. Key Developments"] = 2

        toc_data_levels["*The following companies are listed for indicative purposes only. Similar information will be provided for each, with detailed financial data available exclusively for publicly listed companies.*"] = 1

        for i, name in enumerate(company_list[1:], start=2):
            toc_data_levels[f"{x+2}.{i}. {name}"] = 1

        toc_data_levels[f"{x+3}. Conclusion & Recommendation"] = 0

        table_taxonomy = {
            f"BY {key.upper()}": list(value.keys()) for key, value in nested_dict.items()
        }

        # --- Slide Data Dictionary ---
        slide_data = {
            0: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
                "context": context,
            },
            1: {
                "heading_2": f"{headline}({currency.upper()})",
                "hyear": f"Historical year - {historical_year}",
                "fyear": f"Forecast year - {forecast_period}",
                "byear": f"Base year - {base_year}",
            }, 
            2: {
                "heading_2": f"{headline}({currency.upper()})",
                "hyear": f"Historical year - {historical_year}",
                "fyear": f"Forecast year - {forecast_period}",
                "byear": f"Base year - {base_year}",
            },
            3: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            8: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            10: {
                "org_1": org_1,
                "org_2": org_2,
                "org_3": org_3,
                "org_4": org_4,
                "org_5": org_5,
                "paper_1": paper_1,
                "paper_2": paper_2,
                "paper_3": paper_3,
                "paper_4": paper_4,
                "paper_5": paper_5,
            },
            12: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            13: {
                "heading_2": f"{headline}, {currency}",
                "mpara": mpara_11,
                "para": para_11,
                "amount_1": rev_current,
                "amount_2": rev_future,
            },
            14: {
                "heading": headline,
                "amount_1": f"{rev_current} {value_in}",
                "amount_2": f"{rev_future} {value_in}",
            },
            15: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            16: {"heading": headline, "para": para_14},
            17: {"para": para_15},
            19: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            21: {
                "heading": headline,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
                "types": s_segment,
            },
            22: {
                "heading": headline,
                "type_1": main_topic[0] if main_topic else "Type 1",
                "timeline": historical_year,
                "cur": f"{cur} {value_in}",
            },
            23: {
                "2_heading": headline_2,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            24: {"heading": headline, "timeline": historical_year, "cur": f"{cur} {value_in}"},
            25: {
                "2_heading": headline_2,
                "type_1": main_topic[0] if main_topic else "Type 1",
                "timeline": historical_year,
                "cur": f"{cur} {value_in}",
            },
            26: {
                "2_heading": headline_2,
                "type_1": main_topic[0] if main_topic else "Type 1",
                "timeline": historical_year,
                "cur": f"{cur} {value_in}",
            },
            27: {
                "2_heading": headline_2,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            28: {"heading": headline},
            29: {
                "company": company_info["company_name"],
                "e": company_info["employee_count"],
                "rev": str(revenue_list[1]),
                "h": company_info["headquarters"],
                "geo": company_info["geographic_presence"],
                "es": company_info["estd"],
                "website": company_info["website"],
            },
            30: {
                "2_heading": headline_2,
                "timeline": f"Historic Year {historical_year} and Forecast to {forecast_year}",
            },
            31: {
                "company": company_info["company_name"],
                "e": company_info["employee_count"],
                "rev": str(revenue_list[2]),
                "ownership": company_info["ownership"],
                "h": company_info["headquarters"],
                "website": company_info["website"],
                "es": company_info["estd"],
                "product": company_info["top_product"],
                "para": company_info["short_description_company"],
            },
            32: {"company": company_info["company_name"]},
            33: {"company": company_info["company_name"]},
        }

        # --- PRESENTATION MODIFICATION ---
        print("Loading base presentation 'testppt.pptx'...")
        if not os.path.exists("testppt.pptx"):
            return jsonify({
                'error': 'Template file not found',
                'message': 'Please ensure testppt.pptx is in the project directory'
            }), 500
            
        prs = Presentation("testppt.pptx")

        # Step 1: Replace text in presentation
        for slide in prs.slides:
            data = slide_data.get(prs.slides.index(slide), {})
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for key, value in data.items():
                            token = f"{{{{{key}}}}}"
                            replace_text_in_paragraph(paragraph, token, value)

        # Apply color-preserving text replacement
        for slide in prs.slides:
            data = slide_data.get(prs.slides.index(slide), {})
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for key, value in data.items():
                            token = f"{{{{{key}}}}}"
                            replace_text_preserving_color(paragraph, token, value)

        # Step 2: Add taxonomy boxes to slide 2 (index 1)
        print("Adding taxonomy boxes...")
        generator = TaxonomyBoxGenerator(prs)
        generator.add_taxonomy_boxes(1, table_taxonomy)

        # Step 3: Perform text replacements in tables on specific slides
        print("Performing text replacements inside tables...")
        table_slide_indices = [10, 16, 17, 22, 24, 25, 26, 29, 31, 32, 33]
        replace_text_in_tables(prs, table_slide_indices, slide_data)

        # Step 4: Add and populate the Table of Contents slides
        print("Creating Table of Contents...")
        toc_slide_indices = [4, 5, 6, 7]
        add_toc_to_slides(prs, toc_data_levels, toc_slide_indices)

        # Step 5: Add tables and charts
        target_slide_indices = [22, 25, 26]
        graph_table = list(nested_dict[main_topic[0]].keys()) if main_topic else []
        total_rows = len(graph_table)
        
        # Logic for row labels
        max_visible_types = 3
        row_labels = [graph_table[i] for i in range(min(total_rows, max_visible_types))]
        if total_rows > max_visible_types:
            row_labels.append("Others")
        row_labels.append("Total")

        # Table columns
        years = [str(y) for y in range(2019, 2033)]
        columns = [""] + years + ["CAGR (2025–2032)"]
        num_rows = len(row_labels) + 1
        num_cols = len(columns)

        # Color definitions
        header_rgb = RGBColor(49, 6, 126)
        border_rgb = RGBColor(166, 166, 166)
        alt_row_colors = [RGBColor(231, 231, 231), RGBColor(255, 255, 255)]

        # Font family mappings
        font_mapping = {
            "header": "Poppins",
            "first_col": "Poppins Bold",
            "values": "Poppins Medium",
        }

        # Loop through slides for tables
        for slide_index in target_slide_indices:
            if slide_index < len(prs.slides):
                slide = prs.slides[slide_index]

                # Table placement
                left = Inches(0.4)
                top = Inches(4.05)
                width = Inches(8)
                height = Inches(0.72 + num_rows * 0.3)
                table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table

                # Populate header row
                for col_index, header in enumerate(columns):
                    cell = table.cell(0, col_index)
                    cell.text = header.replace("\n", " ").strip()
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = header_rgb

                    para = cell.text_frame.paragraphs[0]
                    para.alignment = PP_ALIGN.CENTER
                    run = para.runs[0] if para.runs else para.add_run()
                    if col_index != num_cols - 1:
                        run.font.size = Pt(5.7)
                    else:
                        run.font.size = Pt(8)
                    if col_index != num_cols - 1:
                        cell.text_frame.word_wrap = False
                    cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.name = font_mapping["header"]

                    set_cell_border(cell)

                # Populate data rows
                for row_index, label in enumerate(row_labels, start=1):
                    row_color = alt_row_colors[(row_index - 1) % 2]

                    for col_index in range(num_cols):
                        cell = table.cell(row_index, col_index)

                        # Fill content
                        if col_index == 0:
                            cell.text = label
                        elif col_index == num_cols - 1:
                            cell.text = "XX%"
                        else:
                            cell.text = "XX"

                        para = cell.text_frame.paragraphs[0]
                        para.alignment = PP_ALIGN.CENTER
                        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                        if col_index == 0:
                            para.font.size = Pt(8)
                            para.font.name = font_mapping["first_col"]
                            para.font.bold = True
                        else:
                            para.font.size = Pt(9)
                            para.font.name = font_mapping["values"]

                            if label == "Total" and col_index == num_cols - 1:
                                para.font.bold = True
                            if row_index == num_rows - 1:
                                para.font.bold = True

                        cell.fill.solid()
                        cell.fill.fore_color.rgb = row_color
                        set_cell_border(cell)

                # Column widths
                for col_index in range(num_cols):
                    if col_index == 0:
                        table.columns[col_index].width = Inches(1)
                    elif col_index == num_cols - 1:
                        table.columns[col_index].width = Inches(0.8)
                    else:
                        table.columns[col_index].width = Inches(0.4)

        # Add charts to slides
        if main_topic:
            # Determine Columns
            if total_rows <= max_visible_types:
                chart_columns = graph_table
            else:
                chart_columns = graph_table[:max_visible_types] + ["Others"]

            # Insert Chart in Each Slide
            for idx in target_slide_indices:
                if idx < len(prs.slides):
                    slide = prs.slides[idx]
                    data = generate_random_data()
                    
                    # Create chart
                    create_chart_on_slide(
                        slide, data, chart_columns,
                        Inches(0.4), Inches(1.1), Inches(12.5), Inches(2.8)
                    )

        # Add revenue chart to slide 31
        slide_index = 32
        if slide_index < len(prs.slides):
            slide = prs.slides[slide_index]

            # Set the chart data
            chart_data = CategoryChartData()
            chart_data.categories = ["2022", "2023", "2024"]
            chart_data.add_series("Revenue", revenue_list)

            # Define position and size (in inches)
            x = Inches(1)
            y = Inches(1.3)
            cx = Inches(5.6)
            cy = Inches(2.8)

            # Add the clustered column chart
            chart = slide.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
            ).chart
            chart.plots[0].gap_width = 50
            
            # Optional: Remove chart title
            chart.has_title = False

            # Optional: Customize legend
            chart.has_legend = True
            chart.legend.font.size = Pt(10)
            chart.legend.font.name = "Poppins"
            chart.legend.font.bold = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False

            value_axis = chart.value_axis
            value_axis.visible = False
            value_axis.has_major_gridlines = False
            value_axis.has_minor_gridlines = False
            value_axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
            value_axis.major_tick_mark = XL_TICK_MARK.NONE
            value_axis.minor_tick_mark = XL_TICK_MARK.NONE
            category_axis = chart.category_axis
            category_axis.major_tick_mark = XL_TICK_MARK.NONE
            category_axis.minor_tick_mark = XL_TICK_MARK.NONE
            cat_axis = chart.category_axis
            cat_axis.tick_labels.font.size = Pt(10)
            cat_axis.tick_labels.font.name = "Poppins"
            cat_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW

            # Optional: Set custom color for bars (dark purple)
            series = chart.series[0]
            fill = series.format.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(49, 9, 126)

        # Save the final presentation
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"report_{timestamp}.pptx"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        print(f"Saving the final presentation to '{filepath}'...")
        prs.save(filepath)
        print("✅ Script finished successfully!")
        
        return jsonify({
            'success': True,
            'filename': filename,
            'message': 'PowerPoint generated successfully'
        })
        
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            'error': 'Failed to generate PowerPoint',
            'message': str(e)
        }), 500


@app.route('/download/<filename>')
def download_file(filename):
    try:
        safe_filename = secure_filename(filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({'error': 'File not found'}), 404


if __name__ == '__main__':
    # Create .env file if it doesn't exist
    if not os.path.exists('.env'):
        with open('.env', 'w') as f:
            f.write('OPENAI_API_KEY=your_openai_api_key_here\n')
            f.write('GEMINI_API_KEY=your_gemini_api_key_here\n')
        print("Created .env file. Please add your API keys.")
    
    # Create HTML template if it doesn't exist
    if not os.path.exists('templates/index.html'):
        os.makedirs('templates', exist_ok=True)
        print("Please save the HTML content from the artifact to templates/index.html")
    
    # Run the Flask app
    app.run(debug=True, port=5000,)