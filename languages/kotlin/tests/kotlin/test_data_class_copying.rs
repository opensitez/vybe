use crate::helpers::run_prints;

#[test]
fn test_data_class_copy_replaces_one_property() {
    let out = run_prints(
        r#"
        data class User(val name: String, val age: Int)
        fun main() {
            val a = User("a", 1)
            val b = a.copy(age = 2)
            println(a.name)
            println(b.name)
            println(b.age)
        }
    "#,
    );
    assert_eq!(out, &["a", "a", "2"]);
}

#[test]
fn test_data_class_copy_replaces_multiple_properties() {
    let out = run_prints(
        r#"
        data class Box(val id: Int, val tag: String, val active: Boolean)
        fun main() {
            val a = Box(1, "x", false)
            val b = a.copy(id = 2, active = true)
            println(a.id)
            println(b.id)
            println(b.tag)
            println(b.active)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "x", "true"]);
}

#[test]
fn test_data_class_copy_with_nested_defaults() {
    let out = run_prints(
        r#"
        data class Inner(val id: Int)
        data class Outer(val inner: Inner, val label: String)

        fun main() {
            val a = Outer(Inner(1), "a")
            val b = a.copy(inner = Inner(2))
            println(a.inner.id)
            println(b.inner.id)
            println(b.label)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "a"]);
}

#[test]
fn test_data_class_component_functions_in_copy_context() {
    let out = run_prints(
        r#"
        data class PairValue(val a: Int, val b: String)
        fun main() {
            val p = PairValue(1, "x")
            val (a, b) = p
            val copy = p.copy(a = a + 1, b = b.uppercase())
            println(a)
            println(b)
            println(copy.a)
            println(copy.b)
        }
    "#,
    );
    assert_eq!(out, &["1", "x", "2", "X"]);
}

#[test]
fn test_data_class_copy_preserves_reference_for_unchanged() {
    let out = run_prints(
        r#"
        data class Node(val value: IntArray)
        fun main() {
            val src = intArrayOf(1, 2)
            val original = Node(src)
            val copied = original.copy()
            println(original.value.contentEquals(copied.value))
            println(original.value === copied.value)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_data_class_copy_to_mutable_variations() {
    let out = run_prints(
        r#"
        data class Counter(val values: MutableList<Int>)
        fun main() {
            val a = Counter(mutableListOf(1, 2))
            val b = a.copy()
            b.values.add(3)
            println(a.values.joinToString(","))
            println(b.values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "1,2,3"]);
}

#[test]
fn test_data_class_copy_zero_changes_is_identity_like() {
    let out = run_prints(
        r#"
        data class Tag(val name: String)
        fun main() {
            val a = Tag("x")
            val b = a.copy()
            println(a)
            println(b)
            println(a == b)
            println(a === b)
        }
    "#,
    );
    assert_eq!(out, &["Tag(name=x)", "Tag(name=x)", "true", "false"]);
}

#[test]
fn test_data_class_with_generic_copy() {
    let out = run_prints(
        r#"
        data class Holder<T>(val value: T)
        fun main() {
            val a = Holder("x")
            val b = a.copy(value = "y")
            println(a.value)
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["x", "y"]);
}

#[test]
fn test_data_class_named_args_with_copy() {
    let out = run_prints(
        r#"
        data class Config(val host: String, val port: Int, val secure: Boolean)
        fun main() {
            val base = Config("localhost", 80, false)
            val secure = base.copy(port = 443, secure = true)
            println(secure.host)
            println(secure.port)
            println(secure.secure)
        }
    "#,
    );
    assert_eq!(out, &["localhost", "443", "true"]);
}

#[test]
fn test_data_class_destructuring_copy_chain() {
    let out = run_prints(
        r#"
        data class Point(val x: Int, val y: Int)
        fun main() {
            val p1 = Point(1, 2)
            val (x, y) = p1.copy(y = 10)
            println(x)
            println(y)
        }
    "#,
    );
    assert_eq!(out, &["1", "10"]);
}

#[test]
fn test_data_class_copy_with_nullable_property() {
    let out = run_prints(
        r#"
        data class NullableBox(val value: String?)
        fun main() {
            val a = NullableBox(null)
            val b = a.copy(value = "x")
            println(a.value == null)
            println(b.value)
        }
    "#,
    );
    assert_eq!(out, &["true", "x"]);
}

#[test]
fn test_data_class_copy_multiple_instances() {
    let out = run_prints(
        r#"
        data class Row(val id: Int)
        fun main() {
            val first = Row(1)
            val second = first.copy(2)
            val third = second.copy(3)
            println(first.id)
            println(second.id)
            println(third.id)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_data_class_copy_in_function_args() {
    let out = run_prints(
        r#"
        data class User(val name: String, val level: Int)

        fun upgrade(user: User): User = user.copy(level = user.level + 1)

        fun main() {
            val user = User("x", 1)
            val next = upgrade(user)
            println(next.name)
            println(next.level)
        }
    "#,
    );
    assert_eq!(out, &["x", "2"]);
}

#[test]
fn test_data_class_copy_in_collections() {
    let out = run_prints(
        r#"
        data class Item(val name: String, val count: Int)
        fun main() {
            val items = listOf(Item("a", 1), Item("b", 2))
            val upgraded = items.map { it.copy(count = it.count + 10) }
            println(upgraded.joinToString("|") { "${'$'}{it.name}:${'$'}{it.count}" })
        }
    "#,
    );
    assert_eq!(out, &["a:11|b:12"]);
}

#[test]
fn test_data_class_copy_default_arguments() {
    let out = run_prints(
        r#"
        data class Person(val name: String, val age: Int)
        fun main() {
            val a = Person("a", 1)
            val b = a.copy(age = 2)
            val c = a.copy(name = "b")
            println(b.name)
            println(b.age)
            println(c.name)
            println(c.age)
        }
    "#,
    );
    assert_eq!(out, &["a", "2", "b", "1"]);
}

#[test]
fn test_data_class_copy_with_computed_field_in_target() {
    let out = run_prints(
        r#"
        data class Value(val base: Int)
        fun main() {
            val v = Value(2)
            val c = v.copy(base = v.base * 3)
            println(c.base)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_data_class_copying_to_string_equality() {
    let out = run_prints(
        r#"
        data class Flag(val enabled: Boolean)
        fun main() {
            val a = Flag(true)
            val b = a.copy(enabled = false)
            println(a.toString())
            println(b.toString())
            println(a != b)
        }
    "#,
    );
    assert_eq!(out, &["Flag(enabled=true)", "Flag(enabled=false)", "true"]);
}

#[test]
fn test_data_class_copy_hashcode_stability() {
    let out = run_prints(
        r#"
        data class Entry(val key: String, val score: Int)
        fun main() {
            val a = Entry("x", 1)
            val b = a.copy()
            println(a.hashCode() == b.hashCode())
            println(a == b)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_data_class_copy_overwrites_component_position() {
    let out = run_prints(
        r#"
        data class Pair(val left: String, val right: String)
        fun main() {
            val base = Pair("a", "b")
            val next = base.copy("x", right = "y")
            println(next.left)
            println(next.right)
        }
    "#,
    );
    assert_eq!(out, &["x", "y"]);
}

#[test]
fn test_data_class_copy_in_loop() {
    let out = run_prints(
        r#"
        data class Tally(val value: Int)
        fun main() {
            val seed = listOf(1, 2, 3)
            val out = seed.fold(Tally(0)) { acc, next -> acc.copy(value = acc.value + next) }
            println(out.value)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_data_class_copy_with_boolean_flip() {
    let out = run_prints(
        r#"
        data class Switch(val on: Boolean)
        fun main() {
            val a = Switch(false)
            val b = a.copy(on = true)
            println(a.on)
            println(b.on)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_data_class_copy_for_secondary_instances() {
    let out = run_prints(
        r#"
        data class Version(val major: Int, val minor: Int, val patch: Int)
        fun bumpPatch(v: Version): Version = v.copy(patch = v.patch + 1)
        fun main() {
            val v = Version(1, 2, 3)
            val n = bumpPatch(v)
            println(n.major)
            println(n.minor)
            println(n.patch)
        }
    "#,
    );
    assert_eq!(out, &["1", "2", "4"]);
}

#[test]
fn test_data_class_copy_works_with_lists() {
    let out = run_prints(
        r#"
        data class Bucket(val items: List<Int>)
        fun main() {
            val base = Bucket(listOf(1, 2))
            val changed = base.copy(items = base.items + listOf(3, 4))
            println(base.items.joinToString(","))
            println(changed.items.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2", "1,2,3,4"]);
}

#[test]
fn test_data_class_copy_with_empty_string() {
    let out = run_prints(
        r#"
        data class Label(val text: String)
        fun main() {
            val a = Label("ok")
            val b = a.copy(text = "")
            println(a.text.isNotEmpty())
            println(b.text.isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_data_class_copy_of_copy() {
    let out = run_prints(
        r#"
        data class Step(val value: Int)
        fun main() {
            val a = Step(1)
            val b = a.copy()
            val c = b.copy(value = b.value + 1)
            println(a.value)
            println(b.value)
            println(c.value)
        }
    "#,
    );
    assert_eq!(out, &["1", "1", "2"]);
}
