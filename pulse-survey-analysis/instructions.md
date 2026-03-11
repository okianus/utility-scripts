# Survey Results Visualization Script

## Input
- Relevant files and folders: under pulse-survey-analysis directory
- File: survey-responses/GOps Engagement Survey – Pulse Check August 2025.xlsx
- Sheet: Form Responses 1
<!-- - Columns: respondent_id, question_1, question_2, rating, comments -->

## Processing
- use the following mapping of the questions to the attributes they're measuring:
  {"Over the past 6 months, I have had opportunities at work to learn and grow.":"Growth and autonomy"
"I have received recognition or praise for doing good work.":"Recognition"
"My fellow employees are committed to doing quality work.":"Meaningful work"
"I have input into the goals or priorities for my team.":"Effective communication"
"If I make a mistake on this team, it is not held against me.":"Team dynamics"
"I know what is expected of me at work.":"Supportive leadership"
"I understand what success looks like for Grafana, and I believe my team is focused on the right, highest-impact work to achieve it.":"Strategic clarity & leverage"
"I have the opportunity to do what I do best every day.":"Meaningful work"
"There is someone at work who encourages my development.":"Team dynamics"
"I have the freedom to decide how to approach my work.":"Growth and autonomy"
"My manager, or someone at work, seems to care about me as a person.":"Recognition"
"I feel empowered to use AI and other tools to increase my impact and help my team move faster.":"Strategic clarity & leverage"
"My manager gives me useful feedback to help me improve.":"Supportive leadership"
"There are open channels for me to share ideas and concerns.":"Effective communication"
}
<!-- - Calculate average rating per question
- Count responses per answer for multiple-choice questions
- Filter out incomplete responses -->

## Output
- Create or update the visualize_survey.py python script that outputs the following:
  1. Overall sentiment breakdown using the following format
    🟩 XX% Positive (Agree / Strongly Agree)
    🟨 YY% Neutral
    🟥 ZZ% Negative (Disagree / Strongly Disagree)
  2. A horizontal bar chart that includes a short description of each question, and shows the negative, neutral and positive percentages, similar to the format shown in support-files/sample-responses-horizontal-graph.png
  3. Responses per squad (AI, Alerting, IRM, SLOs), organized as vertical bar charts, separately for each question, and showing for each squad the negative, neutral and positive percentages. All vertical bar charts should be part of the same file, organized in 3 rows, with as many columns as needed, depending on each question.  
  4. Map each question to the attribute it's measuring using the instructions in Processing section above. Then create a horizontal bar chart that includes the attribute, and shows the negative, neutral and positive percentages, across the entire department (aggregated over all squads). Make sure that an equal number of questions is mapped to each attribute, and report an error if that's not the case.
  5. Now use the same data as in the previous point, but this time present it as a Radar Chart.  
  6. Save charts as pngs in 'output/' folder
  7. Output all the comments that people have included as optional in their survey responses. Use the LLM to gauge the sentiment as positive / neutral / negative, and group them by sentiment. Before each comment, add in brackets the squad of the person.
  8. Analyze the data using the LLM and report 5 things that are working well as well as 5 areas for improvement
  9. Dump all the output created above (the generated pngs, the comments and the analysis of the data) into a PDF titled YEAR-MONTH-gops-survey-summary.pdf in output/ folder.

<!-- - Bar chart of average rating per question
- Pie chart of distribution for each multiple-choice question
- Save charts as PNGs in 'output/' folder -->