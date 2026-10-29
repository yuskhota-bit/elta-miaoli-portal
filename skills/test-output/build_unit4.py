import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

PURPLE = RGBColor(0x70, 0x30, 0xA0)
GREY = RGBColor(0x59, 0x59, 0x59)


def words(lines):
    return len(re.findall(r"[A-Za-z][A-Za-z'\-]*", " ".join(lines)))


# (page, original label, original text, revised lines, illustration, note)
pages = [
    (1, "Pages 1 + 2", "Miao Li Miao is a happy cat. He lives in Miaoli, Taiwan. He hears a loud sound. BOOM! BOOM! BOOM!",
     ["It is the Lantern Festival.", "Red lanterns glow.", "BOOM! BOOM! BOOM!", "“What is that?” asks Miaoli Meow."],
     "Night in a Hakka town in Miaoli, red lanterns along the street. Miaoli Meow stops, ears up, tail swishing, looking toward the sound.",
     "Old pages 1 and 2 merged. “He lives in Miaoli, Taiwan” cut: the picture shows Miaoli. Lantern Festival added because the event takes place then (cultural insight). Children join in “BOOM! BOOM! BOOM!”."),
    (2, "Page 3", "He walks to Yuqing Temple. He sees a big dragon!",
     ["Miaoli Meow runs to the temple.", "Wow! A big dragon!", "The dragon is gold, silver, red, blue and green."],
     "Yuqing Temple front with the long bamboo dragon in the five colours. Miaoli Meow small in the corner, mouth open in wonder.",
     "Temple name moved to the illustration (fewer place names in the text). Five colours added from the cultural insight; colours are taught with the picture. Check the colours against photos of the Miaoli dragon."),
    (3, "Pages 4 + 6", "“What is this?” asks the cat. “This is Dragon Bombing Day!” “The dragon is colorful,” says the cat. “It is made of paper.”",
     ["“What is this?” asks Miaoli Meow.", "“It is a bamboo dragon!” says Grandpa.", "“Look! Its eyes are open.", "The dragon is awake!”"],
     "Grandpa kneels beside Miaoli Meow and points at the dragon’s head. Close view of the open, painted eyes.",
     "“Paper” corrected to “bamboo” (cultural insight). “Colorful” cut: page 2 shows it. Opening the eyes brings the dragon to life; “awake” is easier than “alive” for Pre-A1. One speaker, Grandpa, so children know who explains."),
    (4, "Page 5", "“We are Hakka people,” says a man. “The dragon brings good luck!”",
     ["“We are Hakka people,” says Grandpa.", "“Today is the Dragon Bombing Carnival.", "The dragon brings good luck!”"],
     "Families in everyday winter clothes smiling around the dragon. Grandpa in front with Miaoli Meow.",
     "“A man” becomes Grandpa, so the story has fewer characters. Festival name moved here so page 3 has one idea. Name follows the series guidelines (“Dragon Bombing Carnival”); the insight sheet says “Bombing Dragon Festival”. To confirm."),
    (5, "Pages 7 + 8 (new problem)", "Miao Li Miao wants to help. He helps a little boy. The drums go BOOM! BOOM! The dragon starts to dance!",
     ["The drums go BOOM! BOOM!", "The dragon dances.", "Miaoli Meow is small.", "“Can you see the dragon?” asks Grandpa.", "“No, I cannot!”"],
     "Crowd from behind: tall legs and backs. Miaoli Meow on tiptoe, sad face, only the top of the dragon visible.",
     "The draft had no clear problem (required story element). The problem is now that Miaoli Meow is too small to see. This brings in the sentence frame “Can you see the dragon?” (1st time). “He helps a little boy” moves to page 9 as one helping moment."),
    (6, "Page 9", "It turns left and right. People cheer for the dragon!",
     ["Grandpa lifts Miaoli Meow up high.", "“Can you see the dragon now?”", "“Yes, I can!”", "It turns left. It turns right.", "Dance, dragon, dance!"],
     "Miaoli Meow on Grandpa’s shoulders, big smile. The dragon twists left and right above the crowd.",
     "Solution: Grandpa helps. Frame used a 2nd time. “Dance, dragon, dance!” is a join-in chant; children can sway left and right. Longest page (24 words) because it carries the turn in the story."),
    (7, "Page 10", "Firecrackers go POP! POP! POP! They chase away bad spirits.",
     ["Grown-ups light the firecrackers.", "POP! POP! POP!", "Miaoli Meow covers his ears.", "“The loud sound chases bad luck away!” says Grandpa."],
     "Grown-ups at a safe distance lighting firecrackers near the dragon. Bright sparks. Miaoli Meow and Grandpa standing well back, paws over ears.",
     "Safety: only grown-ups light firecrackers, and the cat stays back. “Bad spirits” softened to “bad luck” for Grades 1–3; a teacher can explain the belief. Cultural check before illustration."),
    (8, "Page 11 (first half)", "“Stand here!” says Miao Li Miao. A sad girl can see now.",
     ["Miaoli Meow looks around.", "A little girl is at the back.", "“Can you see the dragon?”", "“No, I cannot,” she says."],
     "Same crowd-from-behind layout as page 5. A little girl on tiptoe at the back, sad face. Miaoli Meow notices her.",
     "In the draft the girl appears and is helped on the same line. Here she gets her own page, mirroring page 5 so children see the pattern. Frame used a 3rd time."),
    (9, "Page 11 (second half)", "(continued from page 11 of the draft)",
     ["“Stand here!” says Miaoli Meow.", "“Can you see the dragon now?”", "“Yes, I can!”", "The girl smiles."],
     "Miaoli Meow leads the girl to a low step at the front. She sees the dragon and smiles. Grandpa watches proudly.",
     "Keeps the draft’s “Stand here!”. Miaoli Meow now does for the girl what Grandpa did for him, which sets up the theme on page 10. Children answer “Yes, I can!”."),
    (10, "Page 12", "Good luck is helping friends. Miao Li Miao goes home happy.",
     ["At the end, the dragon goes back to the sky.", "“Goodbye, dragon!”", "Good luck is helping friends.", "Miaoli Meow goes home happy."],
     "Smoke rising gently into the night sky above the temple. Miaoli Meow, Grandpa and the girl wave goodbye.",
     "The dragon is sent back to heaven at the end (cultural insight). Show smoke rising, not a close-up of flames. Kept the theme line “Good luck is helping friends.” The link to the next festival is left out until the running order is confirmed."),
]

orig_counts = [words([p[2]]) for p in pages if not p[2].startswith("(")]
draft_pages = [
    "Miao Li Miao is a happy cat. He lives in Miaoli, Taiwan.",
    "He hears a loud sound. BOOM! BOOM! BOOM!",
    "He walks to Yuqing Temple. He sees a big dragon!",
    "\"What is this?\" asks the cat. \"This is Dragon Bombing Day!\"",
    "\"We are Hakka people,\" says a man. \"The dragon brings good luck!\"",
    "\"The dragon is colorful,\" says the cat. \"It is made of paper.\"",
    "Miao Li Miao wants to help. He helps a little boy.",
    "The drums go BOOM! BOOM! The dragon starts to dance!",
    "It turns left and right. People cheer for the dragon!",
    "Firecrackers go POP! POP! POP! They chase away bad spirits.",
    "\"Stand here!\" says Miao Li Miao. A sad girl can see now.",
    "Good luck is helping friends. Miao Li Miao goes home happy.",
]
dc = [words([t]) for t in draft_pages]
rc = [words(p[3]) for p in pages]

doc = Document()
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(10.5)

doc.add_heading("Exploring Miaoli — Book 1 Unit 4", 0)
doc.add_paragraph("October 29th, 2026 UPDATE · Dragon Bombing Carnival story revised in the house style (Baishatun Mazu co-edit, as applied in Unit 3)")
k = doc.add_paragraph("Key: black = original draft · ")
r = k.add_run("purple = suggested edit")
r.font.color.rgb = PURPLE


def table(rows, header):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        c.paragraphs[0].add_run(h).bold = True
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = v
    return t


doc.add_heading("1 · Style comparison", 1)
doc.add_paragraph(
    f"The submitted draft is the opposite problem to Unit 3: its pages are very short ({min(dc)}–{max(dc)} words a page, "
    f"{sum(dc)} in total), below the Book 1 guide of about 20–40, and the story has no clear problem or solution. "
    f"Revised: 10 story pages of {min(rc)}–{max(rc)} words ({sum(rc)} in total). Pages grew only where the story needed it "
    "(the problem, the help, the cultural facts); the rest stays short so the pictures carry the meaning.")
table([
    ("Short pages, one idea each.", f"12 pages of {min(dc)}–{max(dc)} words; some lines jump ideas (page 11: the girl appears and is already helped).", "10 pages, one idea each. Old pages 1+2, 4+6 and 7+8 merged; page 11 split into a problem page and a help page."),
    ("Story elements: setting, mission, problem, solution, ending.", "Setting and ending present; no clear problem; “He helps a little boy” has no reason.", "Problem: Miaoli Meow is too small to see (p.5). Solution: Grandpa lifts him (p.6). He then helps the girl the same way (pp.8–9)."),
    ("One repeated frame children can join in.", "“What is this?” once; no repeated frame.", "“Can you see the dragon?” — “No, I cannot!” / “Yes, I can!” three times (pp.5, 6, 8–9). Carries forward the Unit 1 “I can see” pattern."),
    ("Join-in sounds and actions.", "BOOM and POP already there.", "Kept BOOM and POP; added “Dance, dragon, dance!” with left/right swaying."),
    ("Fewer place names in the text.", "Miaoli, Taiwan and Yuqing Temple in the text.", "Both moved to the illustrations."),
    ("Cultural accuracy.", "Dragon made of paper; no Lantern Festival; no ending for the dragon.", "Bamboo dragon, five colours, eyes opened, Lantern Festival, grown-ups light firecrackers, dragon sent back to the sky. All flagged for review."),
    ("Character name.", "“Miao Li Miao”.", "“Miaoli Meow”, to match Unit 3 (name still pending)."),
], ["What the house style does", "Dragon draft before", "Dragon revised"])

doc.add_heading("2 · Unit information", 1)
table([
    ("Unit Topic:", "Dragon Bombing Carnival"),
    ("Story Outline", "At the Lantern Festival, Miaoli Meow follows the drums to the temple and meets a bamboo dragon. He is too small to see it dance, so Grandpa lifts him up. Then he helps a little girl see too, and learns that good luck is helping friends."),
    ("Learning goals: Vocabulary Words", "dragon, temple, drum, lantern, firecracker, bamboo, gold, red, small, luck"),
    ("Learning goals: Sentence Pattern", "Can you see the dragon? Yes, I can. / No, I cannot."),
    ("Learning goals: Cultural Understanding", "To understand that Hakka families in Miaoli dance a bamboo dragon at the Lantern Festival, with loud firecrackers to chase bad luck away and bring a good year."),
], ["Item", "Detail"])

doc.add_heading("3 · Page Plan", 1)
t = doc.add_table(rows=1, cols=4)
t.style = "Table Grid"
for i, h in enumerate(["Page", "Story Text", "Illustration Ideas", "Note"]):
    t.rows[0].cells[i].paragraphs[0].add_run(h).bold = True
for num, label, orig, rev, ill, note in pages:
    cells = t.add_row().cells
    cells[0].text = str(num)
    p = cells[1].paragraphs[0]
    o = p.add_run(f"{label}: {orig}")
    o.font.color.rgb = GREY
    hp = cells[1].add_paragraph()
    h = hp.add_run(f"Revised ({words(rev)} words):")
    h.bold = True
    h.font.color.rgb = PURPLE
    for line in rev:
        lr = cells[1].add_paragraph().add_run(line)
        lr.font.color.rgb = PURPLE
    cells[2].text = ill
    cells[3].text = note
doc.add_paragraph("Pages 11–12 stay reserved for the language review and activity pages (no stickers). The original text is shown in grey for comparison; the revised text is what goes forward.")

doc.add_heading("4 · For the team", 1)
doc.add_paragraph("Still to confirm", style="Heading 3")
for s in [
    "Mascot name: “Miaoli Meow” (Unit 3), “Miao Li Miao” (this draft) or “Miaoli Cat” (guidelines).",
    "Festival name: “Dragon Bombing Carnival” (guidelines) or “Bombing Dragon Festival” (insight sheet).",
    "Whether page 10 links to the next festival, once the running order is agreed.",
]:
    doc.add_paragraph(s, style="List Bullet")
doc.add_paragraph("Cultural check needed", style="Heading 3")
for s in [
    "p.2 — the dragon’s five colours.",
    "p.3 — opening the eyes to bring the dragon to life.",
    "p.7 — firecracker scene and “bad luck” wording (the belief is about bad spirits).",
    "p.10 — how the dragon is sent back to heaven, and how much fire to show.",
]:
    doc.add_paragraph(s, style="List Bullet")

out = "/home/user/elta-miaoli-portal/skills/test-output/Exploring_Miaoli_Book1_Unit4_Dragon_Page_Plan_29Oct.docx"
doc.save(out)
print("draft words per page:", dc, "total", sum(dc))
print("revised words per page:", rc, "total", sum(rc))
