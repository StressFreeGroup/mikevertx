#!/usr/bin/env python3
"""
Cover Letter Generator for Mike Katz
Creates a professional cover letter addressing the "WHY" question
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_cover_letter():
    doc = Document()

    # Set up margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # ========== HEADER ==========
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER

    name_run = header.add_run('MICHAEL KATZ\n')
    name_run.bold = True
    name_run.font.size = Pt(24)
    name_run.font.color.rgb = RGBColor(0x7F, 0x00, 0x00)

    contact = header.add_run('San Diego, CA | 917.477.9887 | Mike@MikeVertx.com | linkedin.com/in/MikeVertx')
    contact.font.size = Pt(10)
    contact.font.color.rgb = RGBColor(0x75, 0x75, 0x75)

    doc.add_paragraph()

    # ========== DATE ==========
    date_para = doc.add_paragraph('[Date]')
    date_para.paragraph_format.space_after = Pt(12)

    # ========== RECIPIENT ==========
    recipient = doc.add_paragraph()
    recipient.add_run('[Hiring Manager Name]\n')
    recipient.add_run('[Company Name]\n')
    recipient.add_run('[Address]\n')
    recipient.add_run('[City, State ZIP]')
    recipient.paragraph_format.space_after = Pt(12)

    # ========== SALUTATION ==========
    salutation = doc.add_paragraph('Dear [Hiring Manager Name],')
    salutation.paragraph_format.space_after = Pt(12)

    # ========== THE "WHY" OPENER ==========
    # This is the critical paragraph that addresses the elephant in the room
    why_para = doc.add_paragraph()
    why_run = why_para.add_run("Why would someone who built a $5M business be looking at this role?")
    why_run.bold = True
    why_run.italic = True
    why_para.paragraph_format.space_after = Pt(6)

    why_answer = doc.add_paragraph(
        "It's the question on your mind, so let me answer it directly. After four years building "
        "Stress Free Group from a single RV to 118 vehicles and $5M in revenue, I made a strategic "
        "decision to exit the operational grind and return to what I do best: building brands, "
        "scaling campaigns, and creating memorable experiences. Running my own company taught me "
        "invaluable lessons about P&L ownership, team building, and customer experience—but it also "
        "reminded me that my greatest impact comes from applying 25 years of experiential marketing "
        "expertise to challenges bigger than one company."
    )
    why_answer.paragraph_format.space_after = Pt(12)

    # ========== VALUE PROPOSITION ==========
    value_para = doc.add_paragraph(
        "What I bring isn't just experience—it's perspective that only comes from having done every "
        "job in this industry. I started as a campus brand ambassador at Hofstra. I've been the team "
        "lead coordinating 15 people on-site. I've managed accounts, produced events, negotiated "
        "vendor contracts, and built the talent databases that made national campaigns possible. "
        "When I founded Alternative Brand Marketing, I generated $1.5M in agency revenue serving "
        "brands like Rémy Martin, PepsiCo, and Evian. That journey—from handing out samples to "
        "owning the strategy—means I understand what works at every level."
    )
    value_para.paragraph_format.space_after = Pt(12)

    # ========== SPECIFIC ACHIEVEMENTS ==========
    achievements_para = doc.add_paragraph(
        "At Look Media, I was hired for a 6-week project. Within two weeks, they converted me to "
        "a multi-year contract because I immediately delivered results: 40 new clients, $800K in "
        "new revenue, and a proprietary Talent Portal that transformed how they managed 5,000+ "
        "contractors nationwide. At O.N.E. Natural Experience, I was promoted twice in under two "
        "years and contributed directly to PepsiCo's acquisition of the brand. These aren't just "
        "accomplishments—they're patterns. I identify opportunities others miss and execute at a "
        "level that exceeds expectations."
    )
    achievements_para.paragraph_format.space_after = Pt(12)

    # ========== SPECIFIC INTEREST ==========
    interest_para = doc.add_paragraph(
        "I'm drawn to [Company Name] because [specific reason related to company's work, mission, "
        "or a recent initiative]. My background in [relevant area—e.g., spirits marketing, festival "
        "activations, CPG sampling] aligns directly with your [specific project or need]. I'm not "
        "looking for any job—I'm looking for the right opportunity to apply my expertise where it "
        "will have the greatest impact."
    )
    interest_para.paragraph_format.space_after = Pt(12)

    # ========== CLOSING ==========
    closing_para = doc.add_paragraph(
        "I'd welcome the opportunity to discuss how my combination of entrepreneurial grit and "
        "corporate execution can contribute to your team. I'm available at your convenience and "
        "look forward to the conversation."
    )
    closing_para.paragraph_format.space_after = Pt(24)

    # ========== SIGNATURE ==========
    signature = doc.add_paragraph('Best regards,\n\n\n')
    signature.add_run('Michael Katz').bold = True

    # Save the document
    doc.save('/home/claude/mikevertx/files/COVER_LETTER_MICHAEL_KATZ_2026.docx')
    print("Cover letter created successfully!")

    return doc

if __name__ == '__main__':
    create_cover_letter()
