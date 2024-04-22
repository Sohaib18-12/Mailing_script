def create_message(fullname= "", name="", prefix="", company="", position=""):
    with open("Body.html", "w", encoding='utf-8') as Func:

        Func.write(f"""<html>

<body>

<p>Dear {prefix} {name},</p>

<p>I trust this email finds you well. I am writing to express my enthusiasm for {position} position at {company}.</p>

<p>I would like to have an interview with you to introduce myself and have the opportunity to show my competences. I believe that I am a good candidate for your team as a mechatronics engineer with experience in robotics, software development, machine learning and data analysis.</p>

<p>Attached, you'll find my resume and cover letter, which detail my technical skills and experiences. Beyond my analytical capabilities, I bring a vibrant and positive energy to the workplace, fostering a collaborative and dynamic atmosphere.</p>

<p>I am eager to discuss how my unique blend of technical expertise and positive energy can be an asset to your team. Thank you for considering my application, and I look forward to the opportunity to further discuss my qualifications.</p>

<p>Best regards,<br>
{fullname}<br>
<a href="copy your linkedin profile here">LinkedIn Profile</a></p>

</body>

</html>""")
