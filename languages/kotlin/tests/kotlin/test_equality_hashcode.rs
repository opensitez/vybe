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

#[test]
fn test_equality_is_reflexive_and_symmetric_for_data_classes() {
    let out = run_prints(r#"
        data class Point(val x: Int, val y: Int)

        fun main() {
            val first = Point(1, 2)
            val second = Point(1, 2)
            println(first == first)
            println(first == second)
            println(second == first)
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_hashcode_matches_equals_contract_for_class() {
    let out = run_prints(r#"
        data class Item(val id: Int, val label: String)

        fun main() {
            val left = Item(1, "a")
            val right = Item(1, "a")
            val set = setOf(left, right)
            println(set.size)
            println(left.hashCode() == right.hashCode())
        }
    "#);
    assert_eq!(out, &["1", "true"]);
}

#[test]
fn test_data_class_to_string_includes_all_fields() {
    let out = run_prints(r#"
        data class Entry(val id: Int, val label: String, val active: Boolean)

        fun main() {
            println(Entry(2, "x", true).toString())
        }
    "#);
    assert_eq!(out, &["Entry(id=2, label=x, active=true)"]);
}

#[test]
fn test_reference_equality_still_work_for_same_and_different_instances() {
    let out = run_prints(r#"
        class Item(val label: String)

        fun main() {
            val first = Item("x")
            val second = Item("x")
            val same = first
            println(first === first)
            println(first === same)
            println(first === second)
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_array_content_equality_uses_reference_only() {
    let out = run_prints(r#"
        fun main() {
            val left = arrayOf(1, 2, 3)
            val right = arrayOf(1, 2, 3)
            println(left == right)
            println(left.contentToString())
            println(contentDeepToString(arrayOf(left, right)))
        }
    "#);
    assert_eq!(out, &["false", "[1, 2, 3]", "[[1, 2, 3], [1, 2, 3]]"]);
}

#[test]
fn test_data_class_copy_changes_only_targeted_fields() {
    let out = run_prints(r#"
        data class Pair(val left: Int, val right: String)

        fun main() {
            val source = Pair(1, "a")
            val updated = source.copy(left = 3)
            println(source.left)
            println(updated.left)
            println(updated.right)
            println(source.right)
        }
    "#);
    assert_eq!(out, &["1", "3", "a", "a"]);
}

#[test]
fn test_map_lookup_uses_hashcode_and_equals() {
    let out = run_prints(r#"
        data class Key(val id: String)

        fun main() {
            val map = hashMapOf(Key("a") to 1, Key("b") to 2)
            println(map[Key("a")])
            println(map[Key("x")])
        }
    "#);
    assert_eq!(out, &["1", "null"]);
}

#[test]
fn test_mutable_property_in_data_class_affects_equality() {
    let out = run_prints(r#"
        data class Item(var value: Int)

        fun main() {
            val left = Item(1)
            val set = hashSetOf(left)
            left.value = 2
            println(set.contains(Item(1)))
            println(set.contains(Item(2)))
        }
    "#);
    assert_eq!(out, &["false", "false"]);
}

#[test]
fn test_list_indexof_uses_structural_equality() {
    let out = run_prints(r#"
        data class Entry(val value: Int)

        fun main() {
            val list = listOf(Entry(1), Entry(2))
            println(list.indexOf(Entry(1)))
            println(list.indexOf(Entry(3)))
        }
    "#);
    assert_eq!(out, &["0", "-1"]);
}

#[test]
fn test_equals_with_different_types_is_false() {
    let out = run_prints(r#"
        data class Item(val value: Int)

        fun main() {
            val item = Item(1)
            println(item == 1)
            println(item.equals("x"))
        }
    "#);
    assert_eq!(out, &["false", "false"]);
}

#[test]
fn test_list_equality_uses_element_structural_equality() {
    let out = run_prints(r#"
        data class Item(val value: Int, val label: String)

        fun main() {
            val left = listOf(Item(1, "a"), Item(2, "b"))
            val right = listOf(Item(1, "a"), Item(2, "b"))
            println(left == right)
            println(left[0] == right[0])
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_set_deduplicates_equivalent_data_keys() {
    let out = run_prints(r#"
        data class Key(val id: String, val version: Int)

        fun main() {
            val set = hashSetOf(Key("x", 1), Key("x", 1), Key("x", 2))
            println(set.size)
            println(set.contains(Key("x", 2)))
            println(set.contains(Key("x", 3)))
        }
    "#);
    assert_eq!(out, &["2", "true", "false"]);
}

#[test]
fn test_map_with_equivalent_data_key_overwrites_value() {
    let out = run_prints(r#"
        data class Key(val id: String)

        fun main() {
            val map = hashMapOf(Key("a") to 1)
            map[Key("a")] = 9
            println(map.size)
            println(map[Key("a")])
        }
    "#);
    assert_eq!(out, &["1", "9"]);
}

#[test]
fn test_nested_data_class_to_string_shows_inner_data() {
    let out = run_prints(r#"
        data class Inner(val value: String)
        data class Outer(val inner: Inner, val active: Boolean)

        fun main() {
            println(Outer(Inner("x"), true).toString())
        }
    "#);
    assert_eq!(out, &["Outer(inner=Inner(value=x), active=true)"]);
}

#[test]
fn test_data_class_with_array_field_is_reference_equality_for_array_member() {
    let out = run_prints(r#"
        data class Holder(val values: Array<Int>)

        fun main() {
            val first = Holder(arrayOf(1, 2, 3))
            val second = Holder(arrayOf(1, 2, 3))
            val third = Holder(first.values)
            println(first == second)
            println(first == third)
            println(first.values === third.values)
        }
    "#);
    assert_eq!(out, &["false", "false", "true"]);
}

#[test]
fn test_data_class_hashcode_reuses_structural_equality_for_nested_non_array_field() {
    let out = run_prints(r#"
        data class Child(val token: String)
        data class Parent(val child: Child)

        fun main() {
            val left = Parent(Child("x"))
            val right = Parent(Child("x"))
            println(left == right)
            println(left.hashCode() == right.hashCode())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_nullable_data_class_comparisons_are_safe() {
    let out = run_prints(r#"
        data class Box(val value: Int?)

        fun main() {
            val empty: Box? = null
            val left: Box? = Box(null)
            val right: Box? = Box(null)
            println(empty == null)
            println(left == right)
            println(left == null)
            println(left === right)
        }
    "#);
    assert_eq!(out, &["true", "true", "false", "false"]);
}

#[test]
fn test_reference_equality_remains_for_non_data_class() {
    let out = run_prints(r#"
        class Item(val id: Int)

        fun main() {
            val first = Item(1)
            val second = first
            val third = Item(1)
            println(first == third)
            println(first === second)
            println(first === third)
        }
    "#);
    assert_eq!(out, &["false", "true", "false"]);
}

#[test]
fn test_map_lookup_uses_hashcode_and_equals_on_data_key_after_mutation_isolated() {
    let out = run_prints(r#"
        data class MutableHolder(var value: Int)

        fun main() {
            val original = MutableHolder(1)
            val map = hashMapOf(original to "start")
            original.value = 2
            println(map.containsKey(MutableHolder(1)))
            println(map.containsKey(MutableHolder(2)))
            println(map[MutableHolder(2)])
        }
    "#);
    assert_eq!(out, &["false", "false", "null"]);
}
