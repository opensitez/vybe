use crate::helpers::run_prints;

#[test]
fn test_property_mutable_field_updates_with_assignment() {
    let out = run_prints(r#"
        class Counter {
            var value: Int = 1
        }

        fun main() {
            val counter = Counter()
            counter.value = 4
            println(counter.value)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_property_primary_constructor_value_access() {
    let out = run_prints(r#"
        class User(val name: String, val age: Int)

        fun main() {
            val user = User("Ari", 27)
            println(user.name)
            println(user.age)
        }
    "#);
    assert_eq!(out, &["Ari", "27"]);
}

#[test]
fn test_property_primary_constructor_mutable_property() {
    let out = run_prints(r#"
        class Box(var item: String)

        fun main() {
            val box = Box("start")
            box.item = "done"
            println(box.item)
        }
    "#);
    assert_eq!(out, &["done"]);
}

#[test]
fn test_property_getter_derived_from_other_property() {
    let out = run_prints(r#"
        class Square(val side: Int) {
            val area: Int
                get() = side * side
        }

        fun main() {
            println(Square(5).area)
        }
    "#);
    assert_eq!(out, &["25"]);
}

#[test]
fn test_property_getter_with_private_backing_var() {
    let out = run_prints(r#"
        class Meter {
            private var raw: Int = 2
            val doubled: Int
                get() = raw * 2
        }

        fun main() {
            println(Meter().doubled)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_property_setter_transforms_input() {
    let out = run_prints(r#"
        class Clamp {
            private var raw: Int = 0
            var value: Int
                get() = raw
                set(next) { raw = next * 10 }
        }

        fun main() {
            val c = Clamp()
            c.value = 3
            println(c.value)
        }
    "#);
    assert_eq!(out, &["30"]);
}

#[test]
fn test_property_setter_validates_negative_values() {
    let out = run_prints(r#"
        class Score {
            private var raw = 0
            var value: Int
                get() = raw
                set(next) { raw = if (next < 0) 0 else next }
        }

        fun main() {
            val score = Score()
            score.value = -4
            println(score.value)
            score.value = 7
            println(score.value)
        }
    "#);
    assert_eq!(out, &["0", "7"]);
}

#[test]
fn test_property_setter_records_previous_state() {
    let out = run_prints(r#"
        class Logarithm {
            private var raw = 0
            var value: Int
                get() = raw
                set(next) { raw = next + 1 }
        }

        fun main() {
            val value = Logarithm()
            value.value = 3
            value.value = 7
            println(value.value)
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_property_override_immutable_readonly_property() {
    let out = run_prints(r#"
        open class Node {
            open val label: String = "base"
        }

        class Leaf : Node() {
            override val label: String = "leaf"
        }

        fun main() {
            val node: Node = Leaf()
            println(node.label)
        }
    "#);
    assert_eq!(out, &["leaf"]);
}

#[test]
fn test_property_override_mutable_readwrite_property() {
    let out = run_prints(r#"
        interface CounterLike {
            var count: Int
        }

        class Stateful : CounterLike {
            private var raw = 1
            override var count: Int
                get() = raw
                set(next) { raw = next + 1 }
        }

        fun main() {
            val c: CounterLike = Stateful()
            c.count = 2
            println(c.count)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_property_overridden_accessor_updates_backing_field() {
    let out = run_prints(r#"
        interface ValueSource {
            var value: Int
        }

        class Wrapper : ValueSource {
            private var raw = 4
            override var value: Int
                get() = raw
                set(next) { raw = next - 2 }
        }

        fun main() {
            val value: ValueSource = Wrapper()
            value.value = 10
            println(value.value)
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_property_companion_object_shared_state() {
    let out = run_prints(r#"
        class Factory {
            companion object {
                var created: Int = 0
            }

            init {
                Factory.created += 1
            }
        }

        fun main() {
            Factory()
            Factory()
            println(Factory.created)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_companion_with_instance_and_class_read() {
    let out = run_prints(r#"
        class Counter {
            companion object {
                var next: Int = 0
            }

            fun take(): Int {
                Counter.next += 1
                return Counter.next
            }
        }

        fun main() {
            val c1 = Counter()
            val c2 = Counter()
            println(c1.take())
            println(c2.take())
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_property_val_reference_still_mutable_object() {
    let out = run_prints(r#"
        class Holder {
            val values = mutableListOf(1)
        }

        fun main() {
            val holder = Holder()
            holder.values.add(2)
            println(holder.values.size)
        }
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_property_initializer_order_with_dependency() {
    let out = run_prints(r#"
        class Grid {
            val width = 3
            val height = 4
            val area = width * height
        }

        fun main() {
            val grid = Grid()
            println(grid.width)
            println(grid.height)
            println(grid.area)
        }
    "#);
    assert_eq!(out, &["3", "4", "12"]);
}

#[test]
fn test_property_getter_uses_private_helper_method() {
    let out = run_prints(r#"
        class Formatter {
            private var raw = "kotlin"
            val formatted: String
                get() = raw.uppercase()
        }

        fun main() {
            val formatter = Formatter()
            println(formatter.formatted)
        }
    "#);
    assert_eq!(out, &["KOTLIN"]);
}

#[test]
fn test_property_nullable_value_default_null() {
    let out = run_prints(r#"
        class Note {
            var text: String? = null
        }

        fun main() {
            val note = Note()
            println(note.text == null)
            note.text = "ok"
            println(note.text)
        }
    "#);
    assert_eq!(out, &["true", "ok"]);
}

#[test]
fn test_property_nullable_with_setter_default_handling() {
    let out = run_prints(r#"
        class Holder {
            private var raw: String? = null
            var value: String?
                get() = raw
                set(next) { raw = next ?: "" }
        }

        fun main() {
            val h = Holder()
            h.value = null
            println("[" + h.value + "]")
        }
    "#);
    assert_eq!(out, &["[]"]);
}

#[test]
fn test_property_accessor_reacts_to_mutating_inputs() {
    let out = run_prints(r#"
        class ScoreBoard {
            var raw = 3
            val rating: Int
                get() = raw * 2
        }

        fun main() {
            val board = ScoreBoard()
            board.raw = 4
            println(board.rating)
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_property_setter_updates_derived_backing_value() {
    let out = run_prints(r#"
        class Range {
            private var current: Int = 0
            var base: Int
                get() = current
                set(next) { current = if (next > 100) 100 else next }
        }

        fun main() {
            val r = Range()
            r.base = 150
            println(r.base)
        }
    "#);
    assert_eq!(out, &["100"]);
}

#[test]
fn test_property_overrides_across_interface_and_class_chain() {
    let out = run_prints(r#"
        interface Named {
            val name: String
        }

        open class Animal : Named {
            override val name: String = "animal"
        }

        class Dog : Animal() {
            override val name: String = "dog"
        }

        fun main() {
            val named: Named = Dog()
            println(named.name)
        }
    "#);
    assert_eq!(out, &["dog"]);
}

#[test]
fn test_property_multiple_instances_are_isolated() {
    let out = run_prints(r#"
        class Tracker {
            var score: Int = 0
        }

        fun main() {
            val a = Tracker()
            val b = Tracker()
            a.score = 4
            b.score = 9
            println(a.score + b.score)
        }
    "#);
    assert_eq!(out, &["13"]);
}

#[test]
fn test_property_in_local_class_scope() {
    let out = run_prints(r#"
        fun makeTag(prefix: String): String {
            class Box {
                val label: String = prefix + "-box"
            }
            return Box().label
        }

        fun main() {
            println(makeTag("new"))
        }
    "#);
    assert_eq!(out, &["new-box"]);
}

#[test]
fn test_top_level_property_read() {
    let out = run_prints(r#"
        val welcome = "hello"

        fun main() {
            println(welcome)
        }
    "#);
    assert_eq!(out, &["hello"]);
}

#[test]
fn test_top_level_property_mutation_and_read() {
    let out = run_prints(r#"
        var score = 0

        fun inc() {
            score += 5
        }

        fun main() {
            println(score)
            inc()
            println(score)
        }
    "#);
    assert_eq!(out, &["0", "5"]);
}

#[test]
fn test_top_level_property_and_function_scope_interaction() {
    let out = run_prints(r#"
        var prefix = "A"

        fun scoped(next: String): String {
            return prefix + ":" + next
        }

        fun main() {
            println(scoped("ok"))
            prefix = "B"
            println(scoped("ok"))
        }
    "#);
    assert_eq!(out, &["A:ok", "B:ok"]);
}

#[test]
fn test_property_same_name_local_and_member_do_not_interfere() {
    let out = run_prints(r#"
        class Holder {
            val value = "member"
        }

        fun main() {
            val value = "local"
            val holder = Holder()
            println(value)
            println(holder.value)
        }
    "#);
    assert_eq!(out, &["local", "member"]);
}
