# vybe-test: python/python_email_message_parsing/test_email_message_walk_generator
# origin: languages/python/tests/python/test_python_email_message_parsing.rs

from email.message import EmailMessage
msg = EmailMessage()
msg.set_content("Text part")
msg.add_alternative("<b>HTML part</b>", subtype="html")
content_types = [p.get_content_type() for p in msg.walk()]
print("text/plain" in content_types)
print("text/html" in content_types)
