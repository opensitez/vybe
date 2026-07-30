use crate::helpers::run_prints;

#[test]
fn test_data_class_constructs_with_field_access() {
    let out = run_prints(r#"
        data class User(val name: String, val age: Int)

        fun main() {
            val a = User("Ada", 30)
            println(a.name)
            println(a.age)
            println(a.toString())
        }
    "#);
    assert_eq!(out, &["Ada", "30", "User(name=Ada, age=30)"]);
}

#[test]
fn test_data_class_copy_keeps_unset_fields() {
    let out = run_prints(r#"
        data class Point(val x: Int, val y: Int)

        fun main() {
            val a = Point(1, 2)
            val b = a.copy()
            println(a == b)
            println(a === b)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_data_class_copy_modifies_single_field() {
    let out = run_prints(r#"
        data class Point(val x: Int, val y: Int)

        fun main() {
            val a = Point(1, 2)
            val b = a.copy(y = 99)
            println(b.x)
            println(b.y)
        }
    "#);
    assert_eq!(out, &["1", "99"]);
}

#[test]
fn test_data_class_named_component_access() {
    let out = run_prints(r#"
        data class PairValue(val left: Int, val right: Int)

        fun main() {
            val p = PairValue(4, 9)
            println(p.component1())
            println(p.component2())
        }
    "#);
    assert_eq!(out, &["4", "9"]);
}

#[test]
fn test_data_class_destructuring_by_index() {
    let out = run_prints(r#"
        data class PairValue(val left: Int, val right: Int)

        fun main() {
            val p = PairValue(7, 11)
            val (left, right) = p
            println(left + right)
        }
    "#);
    assert_eq!(out, &["18"]);
}

#[test]
fn test_data_class_structural_equality_uses_all_fields() {
    let out = run_prints(r#"
        data class Key(val id: Int, val tag: String)

        fun main() {
            val a = Key(1, "x")
            val b = Key(1, "x")
            val c = Key(1, "y")
            println(a == b)
            println(a == c)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_data_class_hash_code_matches_equality() {
    let out = run_prints(r#"
        data class Key(val id: Int, val tag: String)

        fun main() {
            val a = Key(10, "tag")
            val b = Key(10, "tag")
            println(a.hashCode() == b.hashCode())
            println(a.hashCode() != a.hashCode())
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_data_class_with_var_property_supports_mutation() {
    let out = run_prints(r#"
        data class Counter(var value: Int)

        fun main() {
            val c = Counter(3)
            c.value += 4
            println(c.value)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_data_class_copy_updates_mutable_field() {
    let out = run_prints(r#"
        data class Counter(var value: Int)

        fun main() {
            val c = Counter(2)
            val d = c.copy(value = 12)
            println(c.value)
            println(d.value)
            d.value += 1
            println(d.value)
        }
    "#);
    assert_eq!(out, &["2", "12", "13"]);
}

#[test]
fn test_generic_data_class_holds_different_types() {
    let out = run_prints(r#"
        data class Holder<T>(val value: T)

        fun main() {
            val a = Holder(1)
            val b = Holder("x")
            println(a.value)
            println(b.value)
        }
    "#);
    assert_eq!(out, &["1", "x"]);
}

#[test]
fn test_generic_data_class_equality_depends_on_payload() {
    let out = run_prints(r#"
        data class Holder<T>(val value: T)

        fun main() {
            val a = Holder(1)
            val b = Holder(1)
            val c = Holder(2)
            println(a == b)
            println(a == c)
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_data_class_in_set_uses_equality() {
    let out = run_prints(r#"
        data class Item(val id: Int)

        fun main() {
            val first = Item(1)
            val second = Item(1)
            val values = setOf(first)
            println(values.contains(second))
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_data_class_as_map_key_round_trip_lookup() {
    let out = run_prints(r#"
        data class Entry(val k: Int, val v: Int)

        fun main() {
            val map = mapOf(Entry(1, 2) to "ok")
            println(map[Entry(1, 2)])
            println(map[Entry(2, 1)] == null)
        }
    "#);
    assert_eq!(out, &["ok", "true"]);
}

#[test]
fn test_data_class_with_default_values_preserved_in_copy() {
    let out = run_prints(r#"
        data class Settings(val enabled: Boolean = true, val retries: Int = 3)

        fun main() {
            val base = Settings()
            val copy = base.copy(retries = 7)
            println(base.enabled)
            println(base.retries)
            println(copy.enabled)
            println(copy.retries)
        }
    "#);
    assert_eq!(out, &["true", "3", "true", "7"]);
}

#[test]
fn test_data_class_local_scope_declaration() {
    let out = run_prints(r#"
        fun make(): Int {
            data class Local(val value: Int)
            return Local(9).value
        }

        fun main() {
            println(make())
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_data_class_in_list_and_destructure_each() {
    let out = run_prints(r#"
        data class PairNode(val id: Int, val weight: Int)

        fun main() {
            val rows = listOf(PairNode(1, 2), PairNode(3, 4))
            var score = 0
            for ((id, weight) in rows) {
                score += id * weight
            }
            println(score)
        }
    "#);
    assert_eq!(out, &["14"]);
}

#[test]
fn test_data_class_nested_copy_propagates_outer() {
    let out = run_prints(r#"
        data class Child(val value: Int)
        data class Parent(val child: Child, val tag: String)

        fun main() {
            val p1 = Parent(Child(1), "x")
            val p2 = p1.copy(child = Child(9))
            println(p1.child.value)
            println(p2.child.value)
            println(p2.tag)
        }
    "#);
    assert_eq!(out, &["1", "9", "x"]);
}

#[test]
fn test_data_class_when_on_components() {
    let out = run_prints(r#"
        data class Kind(val id: Int, val name: String)

        fun classify(kind: Kind): String {
            val (id, name) = kind
            return if (id == 1) "first:" + name else "other:" + name
        }

        fun main() {
            val first = Kind(1, "root")
            val second = Kind(2, "leaf")
            println(classify(first))
            println(classify(second))
        }
    "#);
    assert_eq!(out, &["first:root", "other:leaf"]);
}

#[test]
fn test_data_class_ordering_by_to_string_is_deterministic() {
    let out = run_prints(r#"
        data class Pair(val a: Int, val b: Int)

        fun main() {
            val a = Pair(2, 1)
            val b = Pair(10, 3)
            val list = listOf(a, b)
            println(list.sortedBy { it.a }.joinToString(";") { it.toString() })
        }
    "#);
    assert_eq!(out, &["Pair(a=2, b=1);Pair(a=10, b=3)"]);
}

#[test]
fn test_data_class_copy_chain_preserves_previous_instances() {
    let out = run_prints(r#"
        data class Node(val id: Int, val label: String)

        fun main() {
            val a = Node(1, "a")
            val b = a.copy(label = "b")
            val c = b.copy(id = 3)
            println(a.label)
            println(b.id)
            println(c.label)
            println(a == b)
            println(b == c)
        }
    "#);
    assert_eq!(out, &["a", "1", "b", "false", "false"]);
}

#[test]
fn test_data_class_with_boolean_and_numeric_fields() {
    let out = run_prints(r#"
        data class Flagged(val enabled: Boolean, val level: Int)

        fun main() {
            val item = Flagged(false, 2)
            println(item.enabled)
            println(item.level)
            val updated = item.copy(enabled = true)
            println(updated.enabled)
            println(updated.level)
        }
    "#);
    assert_eq!(out, &["false", "2", "true", "2"]);
}

#[test]
fn test_data_class_with_nullable_members() {
    let out = run_prints(r#"
        data class Holder(val value: String?)

        fun main() {
            val missing = Holder(null)
            val present = Holder("ok")
            println(missing.value == null)
            println(present.value)
        }
    "#);
    assert_eq!(out, &["true", "ok"]);
}

#[test]
fn test_data_class_plus_operator_style_via_copy() {
    let out = run_prints(r#"
        data class Coord(val x: Int, val y: Int)

        fun main() {
            val origin = Coord(0, 0)
            fun move(point: Coord, dx: Int, dy: Int): Coord {
                return point.copy(x = point.x + dx, y = point.y + dy)
            }
            val moved = move(origin, 3, 4)
            println(moved.x)
            println(moved.y)
            println(origin.x)
        }
    "#);
    assert_eq!(out, &["3", "4", "0"]);
}

#[test]
fn test_data_class_multiple_instances_in_map_lookup_by_copy() {
    let out = run_prints(r#"
        data class Route(val from: Int, val to: Int)

        fun main() {
            val route = Route(1, 2)
            val lookup = mapOf(route to "ok")
            val probe = route.copy(to = 2)
            println(lookup[probe])
        }
    "#);
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_data_class_rebinds_in_iteration() {
    let out = run_prints(r#"
        data class Meter(val id: Int, val value: Int)

        fun main() {
            val items = mutableListOf(Meter(1, 1), Meter(2, 2))
            var sum = 0
            for (item in items) {
                val updated = item.copy(value = item.value + 5)
                sum += updated.value
            }
            println(sum)
            println(items[0].value)
        }
    "#);
    assert_eq!(out, &["12", "1"]);
}

#[test]
fn test_data_class_destructure_with_function_return() {
    let out = run_prints(r#"
        data class Record(val code: Int, val weight: Int)

        fun split(): Record {
            return Record(4, 5)
        }

        fun main() {
            val (code, weight) = split()
            println(code)
            println(weight)
            println(code + weight)
        }
    "#);
    assert_eq!(out, &["4", "5", "9"]);
}

#[test]
fn test_data_class_string_projection_is_stable() {
    let out = run_prints(r#"
        data class Tag(val name: String, val value: Int)

        fun main() {
            val a = Tag("x", 1)
            val b = Tag("y", 1)
            val list = listOf(a, b)
            println(list.joinToString("|") { it.toString() })
        }
    "#);
    assert_eq!(out, &["Tag(name=x, value=1)|Tag(name=y, value=1)"]);
}

#[test]
fn test_data_class_implements_interface_contract() {
    let out = run_prints(r#"
        interface Identifiable { val id: Int }
        data class Item(override val id: Int, val payload: String) : Identifiable

        fun main() {
            val item: Identifiable = Item(7, "payload")
            val a = item.id
            println(a)
            println((item as Item).payload)
        }
    "#);
    assert_eq!(out, &["7", "payload"]);
}

#[test]
fn test_data_class_copy_does_not_mutate_source_instance() {
    let out = run_prints(r#"
        data class Pair(val x: Int, val y: Int)

        fun main() {
            val original = Pair(1, 2)
            val copy = original.copy(y = 9)
            println(original.y)
            println(copy.y)
        }
    "#);
    assert_eq!(out, &["2", "9"]);
}

#[test]
fn test_data_class_deeply_nested_destructuring() {
    let out = run_prints(r#"
        data class Line(val start: Int, val end: Int)
        data class Segment(val a: Line, val b: Line)

        fun main() {
            val seg = Segment(Line(1, 2), Line(3, 4))
            val (left, right) = seg
            println(left.start + right.end)
            val (s, e) = left
            println(s)
            println(e)
        }
    "#);
    assert_eq!(out, &["5", "1", "2"]);
}

#[test]
fn test_data_class_as_generic_type_argument() {
    let out = run_prints(r#"
        data class Box<T>(val value: T)
        data class Holder<T>(val value: Box<T>)

        fun main() {
            val holder = Holder(Box("x"))
            println(holder.value.value)
            val copy = holder.copy(value = holder.value.copy(value = "y"))
            println(holder.value.value)
            println(copy.value.value)
        }
    "#);
    assert_eq!(out, &["x", "x", "y"]);
}

#[test]
fn test_data_class_copy_uses_named_and_positional_semantics_together() {
    let out = run_prints(r#"
        data class Range(val start: Int, val end: Int)

        fun main() {
            val base = Range(1, 10)
            val shifted = Range(0, base.end).copy(start = base.start + 1)
            println(base.toString())
            println(shifted.toString())
        }
    "#);
    assert_eq!(out, &["Range(start=1, end=10)", "Range(start=2, end=10)"]);
}
