use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: CSV + JSON + XML + configparser — reading, writing, parsing, configparser
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_csv_write_and_read() {
    let src = r#"
import csv, io

buf = io.StringIO()
writer = csv.writer(buf)
writer.writerow(["name", "age", "city"])
writer.writerow(["Alice", 30, "London"])
writer.writerow(["Bob", 25, "Paris"])

buf.seek(0)
reader = csv.reader(buf)
for row in reader:
    print(row)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['name', 'age', 'city']",
            "['Alice', '30', 'London']",
            "['Bob', '25', 'Paris']"
        ]
    );
}

#[test]
fn test_py_csv_dictreader_dictwriter() {
    let src = r#"
import csv, io

buf = io.StringIO()
fieldnames = ["product", "price", "qty"]
writer = csv.DictWriter(buf, fieldnames=fieldnames)
writer.writeheader()
writer.writerow({"product": "apple", "price": 1.5, "qty": 100})
writer.writerow({"product": "banana", "price": 0.8, "qty": 200})

buf.seek(0)
reader = csv.DictReader(buf)
for row in reader:
    print(dict(row))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "{'product': 'apple', 'price': '1.5', 'qty': '100'}",
            "{'product': 'banana', 'price': '0.8', 'qty': '200'}"
        ]
    );
}

#[test]
fn test_py_csv_custom_delimiter_and_quoting() {
    let src = r#"
import csv, io

data = [["name", "notes"], ["Alice", "Has, a comma"], ["Bob", 'Has "quotes"']]
buf = io.StringIO()
writer = csv.writer(buf, delimiter="|", quoting=csv.QUOTE_MINIMAL)
for row in data:
    writer.writerow(row)

buf.seek(0)
print(buf.getvalue())
"#;
    let result = run_python(src);
    assert!(result.join("\n").contains("Alice"));
    assert!(result.join("\n").contains("Has, a comma"));
}

#[test]
fn test_py_configparser_basic() {
    let src = r#"
import configparser, io

config_str = """
[database]
host = localhost
port = 5432
name = mydb

[cache]
ttl = 300
max_size = 1000
"""

config = configparser.ConfigParser()
config.read_string(config_str)

print(config["database"]["host"])
print(config.getint("database", "port"))
print(config.getint("cache", "ttl"))
print(config.sections())
"#;
    assert_eq!(
        run_python(src),
        vec!["localhost", "5432", "300", "['database', 'cache']"]
    );
}

#[test]
fn test_py_configparser_defaults_and_fallback() {
    let src = r#"
import configparser

config = configparser.ConfigParser(defaults={"timeout": "30"})
config.read_string("[app]\nname=MyApp")

print(config["app"]["name"])
print(config["app"]["timeout"])  # from defaults
print(config.get("app", "missing", fallback="default_val"))
"#;
    assert_eq!(run_python(src), vec!["MyApp", "30", "default_val"]);
}

#[test]
fn test_py_configparser_write() {
    let src = r#"
import configparser, io

config = configparser.ConfigParser()
config["server"] = {"host": "example.com", "port": "443", "ssl": "true"}
config["logging"] = {"level": "INFO"}

buf = io.StringIO()
config.write(buf)
output = buf.getvalue()
print("server" in output)
print("host = example.com" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_xml_etree_basic_parsing() {
    let src = r#"
import xml.etree.ElementTree as ET

xml_str = """<catalog>
    <book id="1"><title>Python Cookbook</title><price>39.99</price></book>
    <book id="2"><title>Fluent Python</title><price>49.99</price></book>
</catalog>"""

root = ET.fromstring(xml_str)
print(root.tag)

for book in root.findall("book"):
    title = book.find("title").text
    price = float(book.find("price").text)
    book_id = book.get("id")
    print(f"{book_id}: {title} ${price}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "catalog",
            "1: Python Cookbook $39.99",
            "2: Fluent Python $49.99"
        ]
    );
}

#[test]
fn test_py_xml_etree_build_and_write() {
    let src = r#"
import xml.etree.ElementTree as ET

root = ET.Element("config")
db = ET.SubElement(root, "database")
ET.SubElement(db, "host").text = "localhost"
ET.SubElement(db, "port").text = "5432"

tree = ET.ElementTree(root)
import io
buf = io.BytesIO()
tree.write(buf, encoding="utf-8", xml_declaration=True)

output = buf.getvalue().decode()
print("localhost" in output)
print("5432" in output)
print("database" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_xml_etree_xpath_find() {
    let src = r#"
import xml.etree.ElementTree as ET

xml_str = """<root>
    <item type="fruit">apple</item>
    <item type="veg">carrot</item>
    <item type="fruit">banana</item>
</root>"""

root = ET.fromstring(xml_str)
fruits = root.findall(".//item[@type='fruit']")
print([el.text for el in fruits])
"#;
    assert_eq!(run_python(src), vec!["['apple', 'banana']"]);
}

#[test]
fn test_py_json_vs_configparser_choice() {
    let src = r#"
import json, configparser

# JSON preserves types
json_config = json.loads('{"port": 8080, "debug": true, "tags": ["a", "b"]}')
print(type(json_config["port"]).__name__)
print(type(json_config["debug"]).__name__)
print(json_config["tags"])

# ConfigParser keeps everything as strings
cfg = configparser.ConfigParser()
cfg.read_string("[app]\nport=8080\ndebug=true")
print(type(cfg["app"]["port"]).__name__)
print(cfg.getboolean("app", "debug"))
"#;
    assert_eq!(
        run_python(src),
        vec!["int", "bool", "['a', 'b']", "str", "True"]
    );
}
