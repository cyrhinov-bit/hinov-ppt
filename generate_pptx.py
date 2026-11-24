import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

def create_presentation():
    prs = Presentation()
    # 16:9 Widescreen standard
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Palette HINOV GROUP — couleurs extraites du logo (assets/logo.png)
    COLOR_BG = RGBColor(6, 13, 31)       # #060D1F Deep Navy
    COLOR_CARD = RGBColor(17, 34, 64)     # #112240 Card Navy
    COLOR_GREEN = RGBColor(74, 208, 123)  # #4AD07B Hinov Green (logo)
    COLOR_EMERALD = RGBColor(16, 185, 129) # #10B981
    COLOR_WHITE = RGBColor(255, 255, 255) # Pure White
    COLOR_MUTED = RGBColor(203, 213, 225) # Slate text
    COLOR_BORDER = RGBColor(56, 189, 248) # Cyan border

    logo_hinov_path = os.path.abspath("assets/logo.png")

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG
        bg.line.fill.background()
        return bg

    def add_header(slide, slide_num, slide_title, category="HINOV GROUP"):
        set_slide_background(slide)

        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.1))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = RGBColor(10, 25, 47)
        hdr.line.color.rgb = RGBColor(30, 58, 100)

        if os.path.exists(logo_hinov_path):
            slide.shapes.add_picture(logo_hinov_path, Inches(0.4), Inches(0.18), height=Inches(0.78))

        txBox = slide.shapes.add_textbox(Inches(2.5), Inches(0.1), Inches(8.333), Inches(0.95))
        tf = txBox.text_frame
        tf.word_wrap = True
        
        p1 = tf.paragraphs[0]
        p1.text = f"{slide_num:02d}  •  {category.upper()}"
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_GREEN
        p1.alignment = PP_ALIGN.CENTER

        p2 = tf.add_paragraph()
        p2.text = slide_title.upper()
        p2.font.size = Pt(17)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_WHITE
        p2.alignment = PP_ALIGN.CENTER

        ftr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(6.9), Inches(13.333), Inches(0.6))
        ftr.fill.solid()
        ftr.fill.fore_color.rgb = RGBColor(10, 25, 47)
        ftr.line.color.rgb = RGBColor(30, 58, 100)

        ftr_tx = slide.shapes.add_textbox(Inches(0.5), Inches(6.95), Inches(12.333), Inches(0.5))
        ftr_tf = ftr_tx.text_frame
        p_ftr = ftr_tf.paragraphs[0]
        p_ftr.text = "HINOV GROUP — Abidjan, Côte d'Ivoire"
        p_ftr.font.size = Pt(10)
        p_ftr.font.bold = True
        p_ftr.font.color.rgb = COLOR_MUTED
        p_ftr.alignment = PP_ALIGN.CENTER

    def add_card(slide, left, top, width, height, title, subtitle, icon="🔹", accent_color=COLOR_GREEN):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = RGBColor(40, 80, 130)
        card.line.width = Pt(1.5)

        txBox = slide.shapes.add_textbox(left + Inches(0.15), top + Inches(0.15), width - Inches(0.3), height - Inches(0.3))
        tf = txBox.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = f"{icon}  {title}"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = accent_color
        
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(11)
        p2.font.color.rgb = COLOR_MUTED
        p2.space_before = Pt(8)
        return card

    def _add_entrance_fade_animations(slide, shape_ids, delays_ms):
        """Ajoute des animations d'entrée (fade) séquentielles sur une liste de formes."""
        ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'

        timing = etree.Element(qn('p:timing'))
        tnLst = etree.SubElement(timing, qn('p:tnLst'))
        par_root = etree.SubElement(tnLst, qn('p:par'))
        cTn_root = etree.SubElement(par_root, qn('p:cTn'))
        cTn_root.set('id', '1')
        cTn_root.set('dur', 'indefinite')
        cTn_root.set('restart', 'never')
        cTn_root.set('nodeType', 'tmRoot')
        childTnLst_root = etree.SubElement(cTn_root, qn('p:childTnLst'))
        seq = etree.SubElement(childTnLst_root, qn('p:seq'))
        seq.set('concurrent', '1')
        seq.set('nextAc', 'seek')
        cTn_main = etree.SubElement(seq, qn('p:cTn'))
        cTn_main.set('id', '2')
        cTn_main.set('dur', 'indefinite')
        cTn_main.set('nodeType', 'mainSeq')
        childTnLst_main = etree.SubElement(cTn_main, qn('p:childTnLst'))

        cid = 3
        for spid, delay in zip(shape_ids, delays_ms):
            # --- État initial : masquer la forme (visibility = hidden) ---
            par_hide = etree.SubElement(childTnLst_main, qn('p:par'))
            cTn_h = etree.SubElement(par_hide, qn('p:cTn'))
            cTn_h.set('id', str(cid)); cid += 1
            cTn_h.set('fill', 'hold')
            stCondLst_h = etree.SubElement(cTn_h, qn('p:stCondLst'))
            cond_h = etree.SubElement(stCondLst_h, qn('p:cond'))
            cond_h.set('delay', '0')
            childTnLst_h = etree.SubElement(cTn_h, qn('p:childTnLst'))
            par_h = etree.SubElement(childTnLst_h, qn('p:par'))
            cTn_h2 = etree.SubElement(par_h, qn('p:cTn'))
            cTn_h2.set('id', str(cid)); cid += 1
            cTn_h2.set('fill', 'hold')
            stCondLst_h2 = etree.SubElement(cTn_h2, qn('p:stCondLst'))
            cond_h2 = etree.SubElement(stCondLst_h2, qn('p:cond'))
            cond_h2.set('delay', '0')
            childTnLst_h2 = etree.SubElement(cTn_h2, qn('p:childTnLst'))
            par_h3 = etree.SubElement(childTnLst_h2, qn('p:par'))
            cTn_h3 = etree.SubElement(par_h3, qn('p:cTn'))
            cTn_h3.set('id', str(cid)); cid += 1
            cTn_h3.set('presetID', '42')
            cTn_h3.set('presetClass', 'entr')
            cTn_h3.set('presetSubtype', '0')
            cTn_h3.set('fill', 'hold')
            cTn_h3.set('grpId', '0')
            cTn_h3.set('nodeType', 'clickEffect')
            stCondLst_h3 = etree.SubElement(cTn_h3, qn('p:stCondLst'))
            cond_h3 = etree.SubElement(stCondLst_h3, qn('p:cond'))
            cond_h3.set('delay', '0')
            childTnLst_h3 = etree.SubElement(cTn_h3, qn('p:childTnLst'))
            set_h = etree.SubElement(childTnLst_h3, qn('p:set'))
            cBhvr_h = etree.SubElement(set_h, qn('p:cBhvr'))
            cTn_hb = etree.SubElement(cBhvr_h, qn('p:cTn'))
            cTn_hb.set('id', str(cid)); cid += 1
            cTn_hb.set('dur', '1')
            cTn_hb.set('fill', 'hold')
            stCondLst_hb = etree.SubElement(cTn_hb, qn('p:stCondLst'))
            cond_hb = etree.SubElement(stCondLst_hb, qn('p:cond'))
            cond_hb.set('delay', '0')
            tgtEl_h = etree.SubElement(cBhvr_h, qn('p:tgtEl'))
            spTgt_h = etree.SubElement(tgtEl_h, qn('p:spTgt'))
            spTgt_h.set('spid', str(spid))
            attrNameLst_h = etree.SubElement(cBhvr_h, qn('p:attrNameLst'))
            attrName_h = etree.SubElement(attrNameLst_h, qn('p:attrName'))
            attrName_h.text = 'style.visibility'
            to_h = etree.SubElement(set_h, qn('p:to'))
            strVal_h = etree.SubElement(to_h, qn('p:strVal'))
            strVal_h.set('val', 'hidden')

            # --- Apparition (fade) au délai donné ---
            par_show = etree.SubElement(childTnLst_main, qn('p:par'))
            cTn_w = etree.SubElement(par_show, qn('p:cTn'))
            cTn_w.set('id', str(cid)); cid += 1
            cTn_w.set('fill', 'hold')
            stCondLst_w = etree.SubElement(cTn_w, qn('p:stCondLst'))
            cond_w = etree.SubElement(stCondLst_w, qn('p:cond'))
            cond_w.set('delay', '0')
            childTnLst_w = etree.SubElement(cTn_w, qn('p:childTnLst'))
            par_d = etree.SubElement(childTnLst_w, qn('p:par'))
            cTn_d = etree.SubElement(par_d, qn('p:cTn'))
            cTn_d.set('id', str(cid)); cid += 1
            cTn_d.set('fill', 'hold')
            stCondLst_d = etree.SubElement(cTn_d, qn('p:stCondLst'))
            cond_d = etree.SubElement(stCondLst_d, qn('p:cond'))
            cond_d.set('delay', str(delay))
            childTnLst_d = etree.SubElement(cTn_d, qn('p:childTnLst'))
            par_fx = etree.SubElement(childTnLst_d, qn('p:par'))
            cTn_fx = etree.SubElement(par_fx, qn('p:cTn'))
            cTn_fx.set('id', str(cid)); cid += 1
            cTn_fx.set('presetID', '42')
            cTn_fx.set('presetClass', 'entr')
            cTn_fx.set('presetSubtype', '0')
            cTn_fx.set('fill', 'hold')
            cTn_fx.set('grpId', '0')
            cTn_fx.set('nodeType', 'clickEffect')
            stCondLst_fx = etree.SubElement(cTn_fx, qn('p:stCondLst'))
            cond_fx = etree.SubElement(stCondLst_fx, qn('p:cond'))
            cond_fx.set('delay', '0')
            childTnLst_fx = etree.SubElement(cTn_fx, qn('p:childTnLst'))
            set_v = etree.SubElement(childTnLst_fx, qn('p:set'))
            cBhvr_v = etree.SubElement(set_v, qn('p:cBhvr'))
            cTn_v = etree.SubElement(cBhvr_v, qn('p:cTn'))
            cTn_v.set('id', str(cid)); cid += 1
            cTn_v.set('dur', '1')
            cTn_v.set('fill', 'hold')
            stCondLst_v = etree.SubElement(cTn_v, qn('p:stCondLst'))
            cond_v = etree.SubElement(stCondLst_v, qn('p:cond'))
            cond_v.set('delay', '0')
            tgtEl_v = etree.SubElement(cBhvr_v, qn('p:tgtEl'))
            spTgt_v = etree.SubElement(tgtEl_v, qn('p:spTgt'))
            spTgt_v.set('spid', str(spid))
            attrNameLst_v = etree.SubElement(cBhvr_v, qn('p:attrNameLst'))
            attrName_v = etree.SubElement(attrNameLst_v, qn('p:attrName'))
            attrName_v.text = 'style.visibility'
            to_v = etree.SubElement(set_v, qn('p:to'))
            strVal_v = etree.SubElement(to_v, qn('p:strVal'))
            strVal_v.set('val', 'visible')
            fade = etree.SubElement(childTnLst_fx, qn('p:fade'))
            cBhvr_f = etree.SubElement(fade, qn('p:cBhvr'))
            cTn_f = etree.SubElement(cBhvr_f, qn('p:cTn'))
            cTn_f.set('id', str(cid)); cid += 1
            cTn_f.set('dur', '1000')
            cTn_f.set('fill', 'hold')
            stCondLst_f = etree.SubElement(cTn_f, qn('p:stCondLst'))
            cond_f = etree.SubElement(stCondLst_f, qn('p:cond'))
            cond_f.set('delay', '0')
            tgtEl_f = etree.SubElement(cBhvr_f, qn('p:tgtEl'))
            spTgt_f = etree.SubElement(tgtEl_f, qn('p:spTgt'))
            spTgt_f.set('spid', str(spid))

        slide.element.append(timing)

    # ==================== SLIDE 1: Page de Garde — HINOV GROUP ====================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    if os.path.exists(logo_hinov_path):
        s1.shapes.add_picture(logo_hinov_path, Inches(0.6), Inches(0.2), height=Inches(0.75))

    tx1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.9), Inches(7.5), Inches(4.3))
    tf1 = tx1.text_frame
    tf1.word_wrap = True

    p_badge = tf1.paragraphs[0]
    p_badge.text = "HINOV GROUP — PRÉSENTATION GÉNÉRALE"
    p_badge.font.size = Pt(11)
    p_badge.font.bold = True
    p_badge.font.color.rgb = COLOR_GREEN
    p_badge.alignment = PP_ALIGN.CENTER

    p_title = tf1.add_paragraph()
    p_title.text = "L'INNOVATION AU SERVICE DE VOTRE PERFORMANCE"
    p_title.font.size = Pt(32)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_WHITE
    p_title.alignment = PP_ALIGN.CENTER
    p_title.space_before = Pt(10)

    p_sub = tf1.add_paragraph()
    p_sub.text = "Solutions en technologies de l'information, infrastructures et communication"
    p_sub.font.size = Pt(18)
    p_sub.font.bold = True
    p_sub.font.color.rgb = COLOR_GREEN
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.space_before = Pt(4)

    p_meta = tf1.add_paragraph()
    p_meta.text = "Abidjan – Côte d'Ivoire | Édition 2026"
    p_meta.font.size = Pt(12)
    p_meta.font.bold = True
    p_meta.font.color.rgb = COLOR_MUTED
    p_meta.alignment = PP_ALIGN.CENTER
    p_meta.space_before = Pt(14)

    p_pres = tf1.add_paragraph()
    p_pres.text = "HINOV GROUP"
    p_pres.font.size = Pt(13)
    p_pres.font.bold = True
    p_pres.font.color.rgb = COLOR_MUTED
    p_pres.alignment = PP_ALIGN.CENTER
    p_pres.space_before = Pt(14)

    p_pres2 = tf1.add_paragraph()
    p_pres2.text = "L'innovation au service de votre performance."
    p_pres2.font.size = Pt(13)
    p_pres2.font.color.rgb = COLOR_MUTED
    p_pres2.alignment = PP_ALIGN.CENTER

    card_r = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.6), Inches(1.9), Inches(3.9), Inches(4.8))
    card_r.fill.solid()
    card_r.fill.fore_color.rgb = COLOR_CARD
    card_r.line.color.rgb = COLOR_GREEN
    card_r.line.width = Pt(2)

    tx_r = s1.shapes.add_textbox(Inches(8.8), Inches(2.2), Inches(3.5), Inches(4.2))
    tf_r = tx_r.text_frame
    tf_r.word_wrap = True
    tf_r.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_rt = tf_r.paragraphs[0]
    p_rt.text = "« L'innovation au service de votre performance. »"
    p_rt.font.size = Pt(14)
    p_rt.font.bold = True
    p_rt.font.color.rgb = COLOR_WHITE
    p_rt.alignment = PP_ALIGN.CENTER
    p_rt.space_before = Pt(10)

    p_rc = tf_r.add_paragraph()
    p_rc.text = "HINOV GROUP accompagne les entreprises, les administrations et les organisations dans leurs besoins en équipements informatiques, infrastructures technologiques et solutions de communication."
    p_rc.font.size = Pt(10)
    p_rc.font.color.rgb = COLOR_MUTED
    p_rc.alignment = PP_ALIGN.CENTER
    p_rc.space_before = Pt(10)

    s1.shapes.add_picture(os.path.abspath("assets/img_hinov_main.jpeg"), Inches(0.8), Inches(6.35), height=Inches(1.15))

    # ==================== SLIDE 2: Notre Vision ====================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, 2, "Notre Vision", "Vision Stratégique")

    b2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.4), Inches(10.9), Inches(1.2))
    b2.fill.solid()
    b2.fill.fore_color.rgb = COLOR_CARD
    b2.line.color.rgb = COLOR_GREEN
    b2.line.width = Pt(2)
    b2_tx = s2.shapes.add_textbox(Inches(1.3), Inches(1.45), Inches(10.7), Inches(0.8))
    b2_tf = b2_tx.text_frame
    b2_tf.word_wrap = True
    p_b2 = b2_tf.paragraphs[0]
    p_b2.text = "Devenir un acteur de référence en Côte d'Ivoire dans les domaines de l'informatique, des infrastructures technologiques, de la communication et des solutions numériques."
    p_b2.font.size = Pt(13)
    p_b2.font.bold = True
    p_b2.font.color.rgb = COLOR_WHITE
    p_b2.alignment = PP_ALIGN.CENTER

    add_card(s2, Inches(1.2), Inches(2.9), Inches(5.3), Inches(1.9), "Notre Ambition", "Devenir un acteur de référence en Côte d'Ivoire dans les domaines de l'informatique, des infrastructures technologiques, de la communication et des solutions numériques.", "🎯", COLOR_GREEN)
    add_card(s2, Inches(6.8), Inches(2.9), Inches(5.3), Inches(1.9), "Notre Contribution", "Contribuer à la modernisation des entreprises en leur apportant des équipements performants, des services de qualité et des solutions innovantes.", "🚀", COLOR_GREEN)

    s2.shapes.add_picture(os.path.abspath("assets/img_vision.jpeg"), Inches(1.2), Inches(5.9), height=Inches(1.2))

    # ==================== SLIDE 3: Notre Mission ====================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, 3, "Notre Mission", "Engagement")

    add_card(s3, Inches(1.2), Inches(1.6), Inches(11.0), Inches(2.6), "Notre Mission", "Notre mission consiste à accompagner les entreprises et les institutions dans la mise en place, l'exploitation et l'amélioration de leur environnement informatique et numérique.", "🎯", COLOR_GREEN)
    add_card(s3, Inches(1.2), Inches(4.8), Inches(11.0), Inches(2.6), "Nos Engagements", "Pour atteindre cet objectif, nous nous engageons à : fournir des équipements de qualité, proposer des solutions adaptées, assurer la maintenance, accompagner la gestion des infrastructures, concevoir et déployer des solutions numériques, développer des solutions de communication et de visibilité digitale.", "🤝", COLOR_GREEN)

    s3.shapes.add_picture(os.path.abspath("assets/img_mission.jpeg"), Inches(0.8), Inches(0.55), height=Inches(0.9))

    # ==================== SLIDE 4: Domaines d'Expertise (carte unique) ====================
    sp = prs.slides.add_slide(blank_layout)
    add_header(sp, 4, "Domaines d'Expertise", "Expertise")
    sp.shapes.add_picture(os.path.abspath("assets/servives.png"), Inches(0.8), Inches(1.2), height=Inches(0.8))
    # Conteneur unique
    container = sp.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(2.2), Inches(10.0), Inches(3.6))
    container.fill.solid()
    container.fill.fore_color.rgb = COLOR_CARD
    container.line.color.rgb = COLOR_BORDER
    container.line.width = Pt(2)
    container.shadow.inherit = False
    # Carte unique
    card = sp.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.7), Inches(2.4), Inches(9.6), Inches(3.2))
    card.fill.solid()
    card.fill.fore_color.rgb = COLOR_CARD
    card.line.color.rgb = COLOR_GREEN
    card.line.width = Pt(1.5)
    card.shadow.inherit = False
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.2)
    tf.margin_bottom = Inches(0.2)

    p_pole = tf.paragraphs[0]
    p_pole.text = "Pôle Service informatique"
    p_pole.font.size = Pt(18)
    p_pole.font.bold = True
    p_pole.font.color.rgb = COLOR_GREEN
    p_pole.alignment = PP_ALIGN.CENTER
    p_pole.space_before = Pt(8)

    items = [
        ("Vente & Fourniture", "Équipements informatiques et bureautiques.", "💻"),
        ("Infogérance", "Disponibilité, sécurité et performance des systèmes.", "🛡️"),
        ("Maintenance", "Diagnostic, réparation et maintenance corrective.", "🔧"),
        ("Câblage IT & Télécom", "Installation d'infrastructures réseau et téléphonique.", "🌐"),
        ("Développement", "Solutions informatiques sur mesure.", "💻"),
        ("Vidéosurveillance", "Fourniture et installation de solutions.", "🛡️"),
        ("Communication & Imprimerie", "Conception graphique, impression et supports.", "🖨️"),
    ]
    for title, sub, icon in items:
        p1 = tf.add_paragraph()
        p1.text = f"{icon}  {title}"
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_GREEN
        p1.alignment = PP_ALIGN.CENTER
        p1.space_before = Pt(6)
        p2 = tf.add_paragraph()
        p2.text = sub
        p2.font.size = Pt(12)
        p2.font.color.rgb = COLOR_MUTED
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(4)

    # ==================== SLIDE 5: Nos Engagements ====================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, 5, "Nos Engagements", "Engagement")

    engagements5 = [
        ("Qualité", "Proposer des équipements et prestations répondant aux exigences.", "✅"),
        ("Réactivité", "Apporter des réponses adaptées aux besoins et urgences.", "⚡"),
        ("Professionnalisme", "Accompagnement sérieux à chaque étape des prestations.", "🤝"),
    ]
    for i, (title, sub, icon) in enumerate(engagements5):
        add_card(s5, Inches(0.8 + i * 4.1), Inches(2.6), Inches(3.6), Inches(3.0), title, sub, icon, COLOR_GREEN)

    add_card(s5, Inches(0.8), Inches(6.0), Inches(3.6), Inches(3.0), "Respect des Engagements", "Respect des délais, des spécifications techniques et des conditions contractuelles.", "🛡️", COLOR_GREEN)
    add_card(s5, Inches(4.9), Inches(6.0), Inches(3.6), Inches(3.0), "Satisfaction Client", "Construire des relations professionnelles durables basées sur la confiance.", "💚", COLOR_GREEN)

    s5.shapes.add_picture(os.path.abspath("assets/img_engagement.jpeg"), Inches(0.8), Inches(6.3), height=Inches(0.9))

    # ==================== SLIDE 6: Clients & Partenaires ====================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, 6, "Clients et Partenaires", "Partenaires")

    card6 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.5))
    card6.fill.solid()
    card6.fill.fore_color.rgb = COLOR_CARD
    card6.line.color.rgb = RGBColor(40, 80, 130)
    card6.line.width = Pt(1.5)

    tx6 = s6.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.3), Inches(4.1))
    tf6 = tx6.text_frame
    tf6.word_wrap = True
    p_t6 = tf6.paragraphs[0]
    p_t6.text = "Une clientèle diversifiée :"
    p_t6.font.size = Pt(12)
    p_t6.font.bold = True
    p_t6.font.color.rgb = COLOR_WHITE
    clients6 = [
        "CORIS BANK CÔTE D'IVOIRE",
        "MEDLOG CÔTE D'IVOIRE",
        "CICT",
        "PNOEV",
        "CNPTIR",
        "MINISTÈRE DES AFFAIRES ÉTRANGÈRES DE CÔTE D'IVOIRE",
        "ÉCOLE ACADÉMIE HORIZON INTERNATIONAL",
        "GROUPE SCOLAIRE MEHI",
        "SMART SECURITY",
        "ETC…",
    ]
    for i, c in enumerate(clients6, start=1):
        p_i = tf6.add_paragraph()
        p_i.text = "•  " + c
        p_i.font.size = Pt(11)
        p_i.font.color.rgb = COLOR_MUTED
        p_i.space_before = Pt(8 if i > 1 else 0)
    p_note6 = tf6.add_paragraph()
    p_note6.text = "Nous accordons une importance particulière à la qualité des relations commerciales et au développement de partenariats durables."
    p_note6.font.size = Pt(10)
    p_note6.font.color.rgb = COLOR_GREEN
    p_note6.space_before = Pt(10)

    s6.shapes.add_picture(os.path.abspath("assets/partner_coris_bank.png"), Inches(0.8), Inches(6.3), height=Inches(0.9))

    # ==================== SLIDE 7: Notre Approche de Travail ====================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, 7, "Notre Approche de Travail", "Approche")

    approach7 = [
        "Analyse et compréhension des besoins du client.",
        "Étude technique et proposition de solutions.",
        "Élaboration d'une offre commerciale adaptée.",
        "Fourniture et déploiement des équipements ou services.",
        "Vérification et contrôle de la qualité des prestations.",
        "Assistance et accompagnement après livraison.",
    ]
    card7 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.6))
    card7.fill.solid()
    card7.fill.fore_color.rgb = COLOR_CARD
    card7.line.color.rgb = RGBColor(40, 80, 130)
    card7.line.width = Pt(1.5)
    tx7 = s7.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.3), Inches(4.4))
    tf7 = tx7.text_frame
    tf7.word_wrap = True
    p_t7 = tf7.paragraphs[0]
    p_t7.text = "Pour chaque projet, HINOV GROUP privilégie une démarche structurée :"
    p_t7.font.size = Pt(12)
    p_t7.font.bold = True
    p_t7.font.color.rgb = COLOR_WHITE
    for i, a in enumerate(approach7, start=1):
        p_i = tf7.add_paragraph()
        p_i.text = "•  " + a
        p_i.font.size = Pt(11)
        p_i.font.color.rgb = COLOR_MUTED
        p_i.space_before = Pt(8 if i > 1 else 0)

    s7.shapes.add_picture(os.path.abspath("assets/img_mission.jpeg"), Inches(0.8), Inches(6.3), height=Inches(0.9))

    # ==================== SLIDE 8: Perspectives de Développement ====================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, 8, "Perspectives de Développement", "Perspectives")

    perspectives8 = [
        "Développement de l'activité de distribution de matériels informatiques.",
        "Renforcement des capacités techniques.",
        "Développement de partenariats avec des fabricants et distributeurs internationaux.",
        "Élargissement du portefeuille de clients institutionnels et professionnels.",
        "Développement de solutions informatiques innovantes.",
        "Renforcement de l'offre de maintenance et d'infogérance.",
    ]
    card8 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.6))
    card8.fill.solid()
    card8.fill.fore_color.rgb = COLOR_CARD
    card8.line.color.rgb = RGBColor(40, 80, 130)
    card8.line.width = Pt(1.5)
    tx8 = s8.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.3), Inches(4.4))
    tf8 = tx8.text_frame
    tf8.word_wrap = True
    p_t8 = tf8.paragraphs[0]
    p_t8.text = "HINOV GROUP ambitionne de renforcer progressivement sa position sur le marché ivoirien à travers :"
    p_t8.font.size = Pt(12)
    p_t8.font.bold = True
    p_t8.font.color.rgb = COLOR_WHITE
    for i, a in enumerate(perspectives8, start=1):
        p_i = tf8.add_paragraph()
        p_i.text = "•  " + a
        p_i.font.size = Pt(11)
        p_i.font.color.rgb = COLOR_MUTED
        p_i.space_before = Pt(8 if i > 1 else 0)

    s8.shapes.add_picture(os.path.abspath("assets/img_vision.jpeg"), Inches(0.8), Inches(1.9), height=Inches(0.9))

    # ==================== SLIDE 9: Pourquoi Choisir HINOV GROUP ? ====================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, 9, "Pourquoi Choisir", "Valeur Ajoutée")

    b9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.4), Inches(10.9), Inches(1.3))
    b9.fill.solid()
    b9.fill.fore_color.rgb = COLOR_CARD
    b9.line.color.rgb = COLOR_GREEN
    b9.line.width = Pt(2)
    b9_tx = s9.shapes.add_textbox(Inches(1.3), Inches(1.45), Inches(10.7), Inches(0.9))
    b9_tf = b9_tx.text_frame
    b9_tf.word_wrap = True
    p_b9 = b9_tf.paragraphs[0]
    p_b9.text = "Choisir HINOV GROUP, c'est faire confiance à un partenaire qui place les besoins de ses clients au centre de ses priorités."
    p_b9.font.size = Pt(13)
    p_b9.font.bold = True
    p_b9.font.color.rgb = COLOR_WHITE
    p_b9.alignment = PP_ALIGN.CENTER

    card9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.9), Inches(11.7), Inches(4.0))
    card9.fill.solid()
    card9.fill.fore_color.rgb = COLOR_CARD
    card9.line.color.rgb = RGBColor(40, 80, 130)
    card9.line.width = Pt(1.5)
    tx9 = s9.shapes.add_textbox(Inches(1.1), Inches(3.1), Inches(11.3), Inches(3.8))
    tf9 = tx9.text_frame
    tf9.word_wrap = True
    p_t9 = tf9.paragraphs[0]
    p_t9.text = "Notre objectif est de contribuer à la performance opérationnelle de nos clients grâce à des solutions adaptées, accessibles et évolutives."
    p_t9.font.size = Pt(12)
    p_t9.font.bold = True
    p_t9.font.color.rgb = COLOR_WHITE
    p_a9 = tf9.add_paragraph()
    p_a9.text = "Nous proposons une approche globale permettant de répondre à plusieurs besoins : fourniture d'équipements, installation, maintenance, assistance informatique, infrastructures réseau et communication."
    p_a9.font.size = Pt(11)
    p_a9.font.color.rgb = COLOR_MUTED
    p_a9.space_before = Pt(8)

    s9.shapes.add_picture(os.path.abspath("assets/img_expertise.jpeg"), Inches(0.8), Inches(1.9), height=Inches(0.9))

    # ==================== SLIDE 10: Nos Coordonnées ====================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, 10, "Nos Coordonnées", "Contact")

    card10 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(11.7), Inches(4.2))
    card10.fill.solid()
    card10.fill.fore_color.rgb = COLOR_CARD
    card10.line.color.rgb = RGBColor(40, 80, 130)
    card10.line.width = Pt(1.5)
    tx10 = s10.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(11.3), Inches(3.9))
    tf10 = tx10.text_frame
    tf10.word_wrap = True
    p_t10 = tf10.paragraphs[0]
    p_t10.text = "HINOV GROUP"
    p_t10.font.size = Pt(16)
    p_t10.font.bold = True
    p_t10.font.color.rgb = COLOR_WHITE
    p_t10.alignment = PP_ALIGN.CENTER
    lines10 = [
        "Siège / implantation : Abidjan, Yopougon, Cité CIE, rue S 259.",
        "Téléphone : (+225) 07 07 02 94 71 / (+225) 07 09 77 03 82",
        "E-mail : hinovgroup@hinovgroup.com",
        "Site internet : www.hinovgroup.com",
        "Localisation : Abidjan – Côte d'Ivoire.",
    ]
    for i, l in enumerate(lines10, start=1):
        p_i = tf10.add_paragraph()
        p_i.text = l
        p_i.font.size = Pt(11)
        p_i.font.color.rgb = COLOR_MUTED
        p_i.alignment = PP_ALIGN.CENTER
        p_i.space_before = Pt(6 if i > 1 else 0)
    p_tag10 = tf10.add_paragraph()
    p_tag10.text = "L'innovation au service de votre performance."
    p_tag10.font.size = Pt(12)
    p_tag10.font.bold = True
    p_tag10.font.color.rgb = COLOR_GREEN
    p_tag10.alignment = PP_ALIGN.CENTER
    p_tag10.space_before = Pt(8)
    p_tag10sub = tf10.add_paragraph()
    p_tag10sub.text = "NOS PARTENAIRES TECHNIQUES"
    p_tag10sub.font.size = Pt(9)
    p_tag10sub.font.bold = True
    p_tag10sub.font.color.rgb = COLOR_MUTED
    p_tag10sub.alignment = PP_ALIGN.CENTER

    s10.shapes.add_picture(os.path.abspath("assets/img_hinov_main.jpeg"), Inches(0.8), Inches(1.9), height=Inches(0.9))

    output_path = "Presentation_Datacenter_Sante_CI.pptx"
    prs.save(output_path)
    print(f"Presentation successfully saved to {output_path}")

if __name__ == "__main__":
    create_presentation()
