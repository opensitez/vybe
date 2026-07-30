use crate::helpers::{compile_ok_check, run_prints};

#[test]
fn test_public_members_are_callable_by_default() {
    let out = run_prints(r#"
        class Item {
            val label = "x"
            fun text(): String = "ok"
        }

        fun main() {
            val item = Item()
            println(item.label)
            println(item.text())
        }
    "#);
    assert_eq!(out, &["x", "ok"]);
}

#[test]
fn test_private_members_are_not_accessible_outside_declaring_class() {
    assert!(!compile_ok_check(r#"
        class Item {
            private val secret: Int = 9
        }

        fun main() {
            val item = Item()
            println(item.secret)
        }
    "#));
}

#[test]
fn test_protected_property_is_visible_to_subclass() {
    let out = run_prints(r#"
        open class Base {
            protected var value: Int = 1
        }

        class Child : Base() {
            fun bump() { value += 2 }
            fun read(): Int = value
        }

        fun main() {
            val child = Child()
            child.bump()
            println(child.read())
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_protected_property_not_visible_to_unrelated_type() {
    assert!(!compile_ok_check(r#"
        open class Base {
            protected var value: Int = 1
        }

        fun leak(base: Base) {
            println(base.value)
        }

        fun main() {
            leak(Base())
        }
    "#));
}

#[test]
fn test_private_constructor_prevents_external_call() {
    assert!(!compile_ok_check(r#"
        class Guard private constructor(val value: Int)

        fun main() {
            val value = Guard(1)
            println(value.value)
        }
    "#));
}

#[test]
fn test_private_class_is_file_local_access_control() {
    let out = run_prints(r#"
        private class Hidden {
            val value = 8
        }

        fun spawn(): Int {
            return Hidden().value
        }

        fun main() {
            println(spawn())
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_private_member_stays_within_class_hierarchy() {
    assert!(!compile_ok_check(r#"
        open class Base {
            private val hidden = 4
        }

        class Child : Base() {
            fun read(): Int = hidden
        }
    "#));
}

#[test]
fn test_internal_property_access_within_module() {
    let out = run_prints(r#"
        class Item {
            internal val value = 9
        }

        fun main() {
            val item = Item()
            println(item.value)
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_override_restricts_private_visibility() {
    assert!(!compile_ok_check(r#"
        open class Base {
            private open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
        }
    "#));
}

#[test]
fn test_private_setter_restricts_external_assignment() {
    assert!(!compile_ok_check(r#"
        class Counter {
            var value: Int = 0
                private set
        }

        fun main() {
            val counter = Counter()
            counter.value = 4
            println(counter.value)
        }
    "#));
}

#[test]
fn test_accessing_private_setter_from_same_class_only() {
    let out = run_prints(r#"
        class Counter {
            var value: Int = 0
                private set

            fun setValue(next: Int) {
                value = next
            }
        }

        fun main() {
            val counter = Counter()
            counter.setValue(11)
            println(counter.value)
        }
    "#);
    assert_eq!(out, &["11"]);
}
