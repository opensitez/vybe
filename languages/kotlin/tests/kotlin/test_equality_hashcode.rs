use crate::helpers::run_prints;

#[test]
fn test_structural_equality_for_data_class() {
    let out = run_prints(r#"
        data class Item(val a: Int, val b: String)

        fun main() {
            val left = Item(1, "x")
            val right = Item(1, "x")
            val other = Item(2, "x")
            println(left == right)
            println(left == other)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_reference_equality_distinguishes_instances() {
    let out = run_prints(r#"
        class Holder(val value: Int)

        fun main() {
            val left = Holder(1)
            val right = Holder(1)
            println(left === right)
            println(left === left)
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_hashcode_matches_structural_equality() {
    let out = run_prints(r#"
        data class Item(val a: Int, val b: String)

        fun main() {
            val left = Item(9, "ok")
            val right = Item(9, "ok")
            println(left.hashCode() == right.hashCode())
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_equals_with_null_and_different_type() {
    let out = run_prints(r#"
        data class Item(val id: Int)

        fun main() {
            val item = Item(1)
            println(item == null)
            println(item == 1)
        }
    "#);
    assert_eq!(out, &["false", "false"]);
}

#[test]
fn test_array_equality_is_reference_based() {
    let out = run_prints(r#"
        fun main() {
            val left = arrayOf(1, 2)
            val right = arrayOf(1, 2)
            println(left == right)
            println(left === right)
        }
    "#);
    assert_eq!(out, &["false", "false"]);
}

#[test]
fn test_to_string_for_data_class_is_stable() {
    let out = run_prints(r#"
        data class Entry(val key: String, val value: Int)

        fun main() {
            println(Entry("a", 3).toString())
        }
    "#);
    assert_eq!(out, &["Entry(key=a, value=3)"]);
}

#[test]
fn test_equality_for_nested_data_classes() {
    let out = run_prints(r#"
        data class Child(val value: Int)
        data class Parent(val child: Child)

        fun main() {
            val a = Parent(Child(1))
            val b = Parent(Child(1))
            println(a == b)
            println(a.child == b.child)
            println(a.child === b.child)
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_mutable_list_component_equality_behavior() {
    let out = run_prints(r#"
        fun main() {
            val left = listOf(1, 2)
            val right = listOf(1, 2)
            println(left == right)
            println(left === right)
            println(left.hashCode() == right.hashCode())
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_classic_primitive_wrappers_stay_by_value() {
    let out = run_prints(r#"
        fun main() {
            val left: Int? = 3
            val right: Int? = 3
            println(left == right)
            println(left === right)
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_equality_after_copy_preserves_original_identity_differences() {
    let out = run_prints(r#"
        data class Point(val x: Int, val y: Int)

        fun main() {
            val base = Point(1, 2)
            val copy = base.copy()
            println(base == copy)
            println(base === copy)
            println(base.x == copy.x)
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_map_key_lookup_uses_equals_contract() {
    let out = run_prints(r#"
        data class Key(val label: String)

        fun main() {
            val cache = mapOf(Key("a") to 1)
            println(cache[Key("a")])
            println(cache.containsKey(Key("x")));
        }
    "#);
    assert_eq!(out, &["1", "false"]);
}

#[test]
fn test_custom_equals_override_controls_contract() {
    let out = run_prints(r#"
        class BadEquals(val value: Int) {
            override fun equals(other: Any?): Boolean {
                if (other !is BadEquals) {
                    return false
                }
                return value == other.value
            }

            override fun hashCode(): Int = value
            override fun toString(): String = "BadEquals(" + value.toString() + ")"
        }

        fun main() {
            val first = BadEquals(2)
            val second = BadEquals(2)
            println(first == second)
            println(first.toString())
            println(first.hashCode())
        }
    "#);
    assert_eq!(out, &["true", "BadEquals(2)", "2"]);
}
