use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: shelve + dbm — persistent dictionary, key-value storage, serialization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_shelve_basic_store_and_retrieve() {
    let src = r#"
import shelve, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    db_path = os.path.join(tmp, "shelf_db")
    with shelve.open(db_path) as db:
        db["user"] = {"name": "Alice", "age": 30}
        db["scores"] = [90, 85, 95]

    with shelve.open(db_path) as db:
        print(db["user"]["name"])
        print(db["scores"])
        print(list(db.keys()))
"#;
    assert_eq!(
        run_python(src),
        vec!["Alice", "[90, 85, 95]", "['user', 'scores']"]
    );
}

#[test]
fn test_py_shelve_writeback_mutation() {
    let src = r#"
import shelve, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    db_path = os.path.join(tmp, "shelf_wb")
    # writeback=True enables mutating mutable objects directly
    with shelve.open(db_path, writeback=True) as db:
        db["items"] = ["a", "b"]
        db["items"].append("c")

    with shelve.open(db_path) as db:
        print(db["items"])
"#;
    assert_eq!(run_python(src), vec!["['a', 'b', 'c']"]);
}

#[test]
fn test_py_shelve_custom_class_persistence() {
    let src = r#"
import shelve, tempfile, os

class Player:
    def __init__(self, name, score):
        self.name = name
        self.score = score

with tempfile.TemporaryDirectory() as tmp:
    db_path = os.path.join(tmp, "shelf_obj")
    with shelve.open(db_path) as db:
        db["p1"] = Player("Bob", 100)

    with shelve.open(db_path) as db:
        p = db["p1"]
        print(p.name, p.score)
"#;
    assert_eq!(run_python(src), vec!["Bob 100"]);
}

#[test]
fn test_py_dbm_basic_byte_storage() {
    let src = r#"
import dbm, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    db_path = os.path.join(tmp, "dbm_store")
    with dbm.open(db_path, "c") as db:
        db[b"key1"] = b"val1"
        db["key2"] = "val2"

    with dbm.open(db_path, "r") as db:
        print(db[b"key1"].decode())
        print(db[b"key2"].decode())
        print(b"key1" in db)
"#;
    assert_eq!(run_python(src), vec!["val1", "val2", "True"]);
}

#[test]
fn test_py_shelve_contains_del_pop() {
    let src = r#"
import shelve, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    db_path = os.path.join(tmp, "shelf_ops")
    with shelve.open(db_path) as db:
        db["a"] = 1
        db["b"] = 2
        print("a" in db)
        print(db.pop("b"))
        del db["a"]
        print("a" in db)
"#;
    assert_eq!(run_python(src), vec!["True", "2", "False"]);
}

#[test]
fn test_py_shelve_get_default() {
    let src = r#"
import shelve, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    db_path = os.path.join(tmp, "shelf_get")
    with shelve.open(db_path) as db:
        db["exists"] = "yes"
        print(db.get("exists"))
        print(db.get("missing", "default_val"))
"#;
    assert_eq!(run_python(src), vec!["yes", "default_val"]);
}
