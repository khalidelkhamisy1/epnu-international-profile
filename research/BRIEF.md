# Shared brief for research subagents (EPNU profile, round 3)

Context: we are building an English PowerPoint profile of **East Port Said National University (EPNU)**,
جامعة شرق بورسعيد الأهلية, New Port Said / East Port Said, Egypt (website https://epnu.edu.eg/en/, founded c. 2024 as a "national/ahleya" university affiliated with Port Said University). Audience: foreign partner universities and delegations (International Relations).

## Program directors (titles as confirmed by the user, use them exactly)
| Faculty | Director | Arabic | Program(s) |
|---|---|---|---|
| Engineering | Assoc. Prof. Mohamed El-Sakka (Mohamed Mohamed Metwally El-Sakka) | محمد محمد متولي السقا | Mechatronics & Automation; Director of Engineering programs; also Director of the IR Office |
| Engineering | Assoc. Prof. Asmaa Rady (Asmaa Rady Ahmed) | أسماء راضي أحمد | AI & Machine Learning Eng.; Construction Eng.; Training & Professional Development Center |
| Computer Science & IT | Prof. Osama El-Baksawy (Osama Mohamed Mostafa El-Baksawy) | أسامة محمد مصطفى البقساوي | Director of CS & IT programs; QA Center |
| Medicine | Assoc. Prof. Sherine El-Khouly (Sherine El-Said El-Khouly) | شيرين السعيد الخولي | Medicine & Surgery (MBBCh) |
| Pharmacy | Assoc. Prof. Mohamed Yasser (Mohamed Yasser Hamdy) | محمد ياسر حمدي | Clinical Pharmacy (PharmD); Entrepreneurship Club |
| Applied Health Sciences | Dr. Salma Khairallah | سلمى خيرالله | Director of the faculty's programs |
| Business Sciences | Dr. Ali Gomaa | علي جمعة | Accounting & Software Applications |

Most of these directors very likely hold their permanent posts at **Port Said University** (PSU, psu.edu.eg) or another Egyptian university; EPNU is PSU's affiliated national university. Check PSU faculty staff pages, Google Scholar, ResearchGate, ORCID, Scopus author pages, Semantic Scholar, OpenAlex, LinkedIn (public snippets only), and EKB journals (journals.ekb.eg).

## Local data you may read
- `D:\AI\EPNU\data\text\en_rendered\pages\<ID>.txt`: English rendered website pages.
  - 10 = Faculties & Programs
  - 1058 = Engineering
  - 3154 = Computer Science & IT
  - 3155 = Medicine
  - 3156 = Pharmacy
  - 1047 = Applied Health Sciences
  - 4176 = Business Sciences
  - 4174 = Arts
  - 4711 = Partnerships
  - 1777 = Academic System
  - 1771 = Admission
  - 1779 = Scholarships
- `D:\AI\EPNU\data\text\en_rendered\posts\*.txt`: English rendered news posts (100).
- `D:\AI\EPNU\data\text\ar\...`: Arabic originals.
- `D:\AI\EPNU\EPNU_International_Relations_Briefing.md`: the current fact base. Do not edit it.
- `D:\AI\EPNU\data\facebook_findings.md`: facts from the official Facebook page.

## Rules
1. **Everything you write must be in English.** Transliterate Arabic names in the forms used above.
2. **Every fact needs a source:** a URL, or a local file path plus page ID. Mark each fact `confidence: high | medium | low`.
3. **Identity disambiguation is critical.** Egyptian names repeat. Accept a profile as the right person only if it matches on at least two of:
   - the exact full name or spelling;
   - affiliation with Port Said University, EPNU or Port Said;
   - a field consistent with the program.

   If you are unsure, report the profile as a "possible match" with your reasoning. Never merge two different people.
4. Do not invent anything. If you find nothing, say "not found" and list where you looked.
5. Read only. Never log in, submit forms, or post anything. Use WebSearch and WebFetch (load them via ToolSearch `select:WebSearch,WebFetch` if they are deferred), plus curl or Python for public APIs.
6. Write your output file to `D:\AI\EPNU\research\` with the exact name given in your task, then finish with a summary of at most 150 words.
