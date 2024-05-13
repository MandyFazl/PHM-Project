# Use spaCy to extract verbs from the second row
        verbs = []
        for cell in rows[1]:
            doc = nlp(cell)
            verbs.extend([token.text for token in doc if token.pos_ == "VERB"])

        # Create nodes for verbs with no border and background at the bottom of the slide
        for i, verb in enumerate(verbs):
            left = i * x_step
            top = slide_height - oval_height - Inches(0.2)
            verb_oval = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, oval_width, oval_height)
            verb_oval.line.fill.solid()
            verb_oval.line.fill.fore_color.rgb = RGBColor(255, 255, 255)  # No border
            verb_oval.line.width = Pt(0)  # No border width
            verb_oval.shadow.inherit = False  # No shadow
            verb_oval.fill.solid()
            verb_oval.fill.fore_color.rgb = RGBColor(255, 255, 255)  # No background
            verb_oval.fill.fore_color.alpha = 0  # 0 alpha for transparency
            verb_text_frame = verb_oval.text_frame
            p = verb_text_frame.add_paragraph()
            p.text = verb
            p.font.size = text_font_size
            p.font.color.rgb = RGBColor(0, 0, 0)  # Black font color
            p.alignment = PP_ALIGN.CENTER