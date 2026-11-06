#!/usr/bin/env python3
"""
Resume Generator for Mike Katz
Creates a professional DOCX resume
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cell_shading(cell, color):
    """Set background color for a table cell."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading)

def create_resume():
    doc = Document()

    # Set up styles
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ========== HEADER ==========
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER

    name_run = header.add_run('MICHAEL KATZ\n')
    name_run.bold = True
    name_run.font.size = Pt(28)
    name_run.font.color.rgb = RGBColor(0x7F, 0x00, 0x00)

    tagline = header.add_run('Marketing & Entertainment Executive | Experiential Marketing Leader\n')
    tagline.font.size = Pt(12)
    tagline.font.color.rgb = RGBColor(0x42, 0x42, 0x42)

    contact = header.add_run('San Diego, CA | 917.477.9887 | Mike@MikeVertx.com | linkedin.com/in/MikeVertx')
    contact.font.size = Pt(10)
    contact.font.color.rgb = RGBColor(0x75, 0x75, 0x75)

    # ========== PROFESSIONAL SUMMARY ==========
    doc.add_paragraph()
    summary_heading = doc.add_paragraph()
    run = summary_heading.add_run('PROFESSIONAL SUMMARY')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

    # Add a line
    summary_heading.paragraph_format.space_after = Pt(6)

    summary = doc.add_paragraph(
        "Results-driven marketing executive with 25+ years of progressive experience building and scaling "
        "experiential marketing operations. Proven track record of climbing every rung of the ladder—from campus "
        "brand ambassador to agency owner—developing deep expertise at each level. Founded and scaled two "
        "businesses: Alternative Brand Marketing ($1.5M+ agency revenue, 50+ brand programs) and Stress Free "
        "Group LLC ($5M+ gross revenue, 118 vehicles, 2,300+ reservations). Known for building scalable systems, "
        "recruiting and developing talent, and delivering measurable ROI across 5,000+ events and 1M+ product samples."
    )
    summary.paragraph_format.space_after = Pt(12)

    # ========== KEY METRICS ==========
    metrics_heading = doc.add_paragraph()
    run = metrics_heading.add_run('KEY METRICS')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

    # Create metrics table
    metrics_table = doc.add_table(rows=1, cols=4)
    metrics_table.autofit = True

    metrics = [
        ('25+', 'Years Experience'),
        ('5,000+', 'Events Produced'),
        ('50+', 'Brand Clients'),
        ('$5M+', 'Business Revenue')
    ]

    row = metrics_table.rows[0]
    for i, (num, label) in enumerate(metrics):
        cell = row.cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        num_run = p.add_run(num + '\n')
        num_run.bold = True
        num_run.font.size = Pt(16)
        num_run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

        label_run = p.add_run(label)
        label_run.font.size = Pt(9)
        label_run.font.color.rgb = RGBColor(0x75, 0x75, 0x75)

    doc.add_paragraph()

    # ========== PROFESSIONAL EXPERIENCE ==========
    exp_heading = doc.add_paragraph()
    run = exp_heading.add_run('PROFESSIONAL EXPERIENCE')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

    # Experience entries
    experiences = [
        {
            'title': 'Owner & President',
            'company': 'Stress Free Group LLC',
            'location': 'San Diego, CA',
            'dates': 'October 2020 – Present',
            'bullets': [
                'Founded and scaled independent RV rental and fleet-management business spanning rentals, repairs, sales, logistics, and large-scale event support',
                'Built vertically integrated operations with direct ownership across P&L, compliance, and customer experience',
                'Grew fleet from 1 vehicle to 118 vehicles at peak capacity with 17 full-time staff',
                'Generated $5M+ gross revenue with 2,300+ customer reservations and industry-leading satisfaction ratings'
            ]
        },
        {
            'title': 'Sr. National Director of Sales & Client Relations',
            'company': 'Look Media USA / The Sampling Agency',
            'location': 'San Diego, CA',
            'dates': 'February 2016 – December 2019',
            'bullets': [
                'Hired for 6-week project; converted to multi-year contract in 2 weeks based on immediate results',
                'Created The Sampling Agency under Look Media umbrella, establishing new revenue vertical',
                'Brought in 40 new clients and $800K revenue increase through strategic business development',
                'Built proprietary Talent Portal software for scheduling, assignments, reporting, and photo uploads',
                'Managed database of 5,000+ contractors across 50 states for national activations'
            ]
        },
        {
            'title': 'Brand Marketing Manager',
            'company': 'Vino Beverages',
            'location': 'New York, NY',
            'dates': 'November 2014 – January 2016',
            'bullets': [
                'Led all marketing for alcohol-removed wine startup targeting health-conscious consumers',
                'Managed quarterly marketing plans with budgets ranging from $30K to $2M',
                'Built full convention calendar strategy generating 1,500+ qualified leads per quarter',
                'Moved 12,000+ units in first 6 months through integrated digital and experiential campaigns'
            ]
        },
        {
            'title': 'Founder',
            'company': 'Alternative Brand Marketing',
            'location': 'New York, NY',
            'dates': 'January 2011 – January 2021',
            'bullets': [
                'Founded full-service experiential marketing agency serving Fortune 500 and startup clients',
                'Led campaign strategy, national sales programming, staffing, and execution for 50+ brand programs',
                'Generated $1.5M+ in agency revenue with 20,000+ contractor database',
                'Delivered multi-market activations for clients including PepsiCo, Rémy Martin, and Evian'
            ]
        },
        {
            'title': 'Regional Field Marketing Manager – East',
            'company': 'O.N.E Natural Experience (PepsiCo)',
            'location': 'Los Angeles / East Coast',
            'dates': 'June 2011 – April 2013',
            'bullets': [
                'Promoted twice from Northeast to full East region; contributed to successful PepsiCo acquisition',
                'Managed 50 direct reports across 600+ events with $100K+ vendor cost savings',
                'Led major festival activations: Bonnaroo, Electric Zoo, Winter Music Conference, EDC'
            ]
        },
        {
            'title': 'Account Producer',
            'company': 'SoHo Experiential',
            'location': 'New York, NY',
            'dates': 'July 2010 – January 2013',
            'bullets': [
                'Managed premium spirits activations for Rémy Martin VSOP, Rémy Martin V, Grand Marnier',
                'Oversaw 50+ brand ambassadors and 500+ events over 3-year engagement',
                'Built nightlife and event activation programs across 20+ major markets'
            ]
        }
    ]

    for exp in experiences:
        # Job title and company line
        job_para = doc.add_paragraph()

        title_run = job_para.add_run(exp['title'])
        title_run.bold = True
        title_run.font.size = Pt(11)

        job_para.add_run(' | ')

        company_run = job_para.add_run(exp['company'])
        company_run.font.color.rgb = RGBColor(0x2E, 0x7D, 0x32)

        job_para.add_run(' | ' + exp['location'])

        # Dates on same line, right-aligned
        job_para.add_run('\t' + exp['dates'])

        job_para.paragraph_format.space_after = Pt(3)

        # Bullets
        for bullet in exp['bullets']:
            bullet_para = doc.add_paragraph(bullet, style='List Bullet')
            bullet_para.paragraph_format.left_indent = Inches(0.25)
            bullet_para.paragraph_format.space_after = Pt(2)

        doc.add_paragraph()

    # ========== SKILLS ==========
    skills_heading = doc.add_paragraph()
    run = skills_heading.add_run('CORE COMPETENCIES')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

    skills = doc.add_paragraph(
        "Experiential Marketing | Event Production | Brand Activation | Sampling Programs | "
        "Staff Recruitment & Training | Multi-Market Campaign Management | P&L Ownership | "
        "Vendor Negotiations | Strategic Partnerships | Business Development | CRM Systems | "
        "Festival Marketing | Nightlife Marketing | Spirits & Beverage | CPG | Entertainment"
    )
    skills.paragraph_format.space_after = Pt(12)

    # ========== EDUCATION ==========
    edu_heading = doc.add_paragraph()
    run = edu_heading.add_run('EDUCATION')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

    edu = doc.add_paragraph()
    edu.add_run('Bachelor of Arts, Speech Communications\n').bold = True
    edu.add_run('Hofstra University | Long Island, NY')

    # ========== CLIENTS ==========
    doc.add_paragraph()
    clients_heading = doc.add_paragraph()
    run = clients_heading.add_run('SELECT CLIENTS')
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0xB7, 0x1C, 0x1C)

    clients = doc.add_paragraph(
        "PepsiCo | Rémy Martin | Grand Marnier | Evian | Danone | Amazon | Audible | DirecTV | "
        "Samsung | Sony | HP | Intel | Qualcomm | Red Bull | Monster Energy | Vita Coco | "
        "Beanfields | Argo Tea | KIND Snacks | Anheuser-Busch | MillerCoors | Bacardi | Diageo"
    )
    clients.paragraph_format.space_after = Pt(6)

    # Save the document
    doc.save('/home/claude/mikevertx/files/RESUME_MICHAEL_KATZ_2026.docx')
    print("Resume created successfully!")

    return doc

if __name__ == '__main__':
    create_resume()
