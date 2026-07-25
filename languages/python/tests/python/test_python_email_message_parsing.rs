use super::helpers::run_python;

// email — EmailMessage, message_from_string, message_from_bytes, add_header, get_content_type, get_payload, add_attachment, is_multipart

#[test]
fn test_email_message_basic_headers() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg["Subject"] = "Test Email"
msg["From"] = "alice@example.com"
msg["To"] = "bob@example.com"
print(msg["Subject"])
print(msg["From"])
print(msg["To"])
"#);
    assert_eq!(out, vec!["Test Email", "alice@example.com", "bob@example.com"]);
}

#[test]
fn test_email_message_set_content_plain_text() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg.set_content("Hello, world!\nThis is a plain text body.")
print(msg.get_content_type())
print(msg.get_content().strip())
"#);
    assert_eq!(out, vec!["text/plain", "Hello, world!\nThis is a plain text body."]);
}

#[test]
fn test_email_message_add_alternative_html() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg.set_content("Plain text version.")
msg.add_alternative("<h1>HTML version</h1>", subtype="html")
print(msg.is_multipart())
print(msg.get_content_type())
"#);
    assert_eq!(out, vec!["True", "multipart/alternative"]);
}

#[test]
fn test_email_message_from_string_parsing() {
    let out = run_python(r#"
import email
raw = """From: sender@domain.com
To: receiver@domain.com
Subject: Meeting Request

Let's meet tomorrow at 10 AM.
"""
msg = email.message_from_string(raw)
print(msg["Subject"])
print(msg["From"])
print(msg.get_payload().strip())
"#);
    assert_eq!(out, vec!["Meeting Request", "sender@domain.com", "Let's meet tomorrow at 10 AM."]);
}

#[test]
fn test_email_message_from_bytes_parsing() {
    let out = run_python(r#"
import email
raw_bytes = b"From: a@b.com\nTo: c@d.com\nSubject: Bytes Test\n\nContent body"
msg = email.message_from_bytes(raw_bytes)
print(msg["Subject"])
print(msg.get_payload().strip())
"#);
    assert_eq!(out, vec!["Bytes Test", "Content body"]);
}

#[test]
fn test_email_message_add_attachment_bytes() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg.set_content("See attached document.")
msg.add_attachment(b"PDF data content here", maintype="application", subtype="pdf", filename="doc.pdf")
print(msg.is_multipart())
parts = list(msg.iter_parts())
print(len(parts))
print(parts[1].get_filename())
"#);
    assert_eq!(out, vec!["True", "2", "doc.pdf"]);
}

#[test]
fn test_email_message_header_case_insensitive() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg["Content-Type"] = "text/plain; charset=utf-8"
print(msg["content-type"])
print(msg["CONTENT-TYPE"])
"#);
    assert_eq!(out, vec!["text/plain; charset=utf-8", "text/plain; charset=utf-8"]);
}

#[test]
fn test_email_message_get_all_multiple_headers() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg.add_header("Received", "from server1")
msg.add_header("Received", "from server2")
received = msg.get_all("Received")
print(received)
"#);
    assert_eq!(out, vec!["['from server1', 'from server2']"]);
}

#[test]
fn test_email_message_replace_header() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg["Subject"] = "Original"
msg.replace_header("Subject", "Updated")
print(msg["Subject"])
"#);
    assert_eq!(out, vec!["Updated"]);
}

#[test]
fn test_email_message_del_header() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg["X-Custom"] = "value"
del msg["X-Custom"]
print("X-Custom" in msg)
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_email_utils_parseaddr() {
    let out = run_python(r#"
from email.utils import parseaddr
name, addr = parseaddr("John Doe <john@example.com>")
print(name)
print(addr)
"#);
    assert_eq!(out, vec!["John Doe", "john@example.com"]);
}

#[test]
fn test_email_utils_formataddr() {
    let out = run_python(r#"
from email.utils import formataddr
formatted = formataddr(("Jane Doe", "jane@example.com"))
print(formatted)
"#);
    assert_eq!(out, vec!["Jane Doe <jane@example.com>"]);
}

#[test]
fn test_email_utils_formatdate() {
    let out = run_python(r#"
from email.utils import formatdate
date_str = formatdate(usegmt=True)
print(isinstance(date_str, str))
print("GMT" in date_str)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_email_utils_parsedate_to_datetime() {
    let out = run_python(r#"
from email.utils import parsedate_to_datetime
dt = parsedate_to_datetime("Mon, 25 Dec 2023 12:00:00 +0000")
print(dt.year, dt.month, dt.day)
print(dt.hour, dt.minute)
"#);
    assert_eq!(out, vec!["2023 12 25", "12 0"]);
}

#[test]
fn test_email_header_decode_header() {
    let out = run_python(r#"
from email.header import decode_header
header_val = "=?utf-8?q?hello_world?="
decoded = decode_header(header_val)
print(decoded[0][0])
print(decoded[0][1])
"#);
    assert_eq!(out, vec!["b'hello world'", "utf-8"]);
}

#[test]
fn test_email_header_make_header() {
    let out = run_python(r#"
from email.header import make_header, decode_header
h = make_header([("café", "utf-8")])
print(str(h))
"#);
    assert_eq!(out, vec!["=?utf-8?q?caf=C3=A9?="]);
}

#[test]
fn test_email_policy_default_policy() {
    let out = run_python(r#"
from email.policy import default
from email import message_from_string
raw = "Subject: =?utf-8?q?Hello_World?=\n\nBody text"
msg = message_from_string(raw, policy=default)
print(msg["Subject"])
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn test_email_message_walk_generator() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg.set_content("Text part")
msg.add_alternative("<b>HTML part</b>", subtype="html")
content_types = [p.get_content_type() for p in msg.walk()]
print("text/plain" in content_types)
print("text/html" in content_types)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_email_message_as_string_serialization() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg["Subject"] = "Test"
msg.set_content("Body")
s = msg.as_string()
print("Subject: Test" in s)
print("Body" in s)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_email_message_as_bytes_serialization() {
    let out = run_python(r#"
from email.message import EmailMessage
msg = EmailMessage()
msg["Subject"] = "Test Bytes"
msg.set_content("Body Bytes")
b = msg.as_bytes()
print(b"Subject: Test Bytes" in b)
"#);
    assert_eq!(out, vec!["True"]);
}
