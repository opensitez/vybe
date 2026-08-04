use crate::helpers::run_prints;

#[test]
fn test_local_class_basic() {
    let out = run_prints(
        r#"
        fun main() {
            class Local(val v: Int)
            println(Local(1).v)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_local_class_with_function() {
    let out = run_prints(
        r#"
        fun main() {
            class C {
                fun id(v: Int): Int = v + 1
            }
            println(C().id(4))
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_data_class() {
    let out = run_prints(
        r#"
        fun main() {
            data class Point(val x: Int, val y: Int)
            val p = Point(1, 2)
            println(p.x + p.y)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_local_class_in_function_scope() {
    let out = run_prints(
        r#"
        fun wrap(v: Int): Int {
            class Local {
                fun value() = v + 1
            }
            return Local().value()
        }
        fun main() {
            println(wrap(3))
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_local_class_in_generic_function() {
    let out = run_prints(
        r#"
        fun <T> use(v: T): String {
            class Holder {
                fun text() = v.toString()
            }
            return Holder().text()
        }
        fun main() {
            println(use(7))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_local_sealed_like_not_allowed() {
    let out = run_prints(
        r#"
        fun main() {
            open class Local(val v: Int)
            class Derived : Local(2)
            println(Derived().v)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_local_class_with_companion_in_function() {
    let out = run_prints(
        r#"
        fun main() {
            class Factory {
                companion object {
                    fun make(v: Int) = Holder(v)
                }
            }
            class Holder(val v: Int)
            println(Factory.make(9).v)
        }
    "#,
    );
    assert_eq!(out, &["9"]);
}

#[test]
fn test_local_enum_like_simple() {
    let out = run_prints(
        r#"
        fun main() {
            enum class Mode { A, B, C }
            println(Mode.B.name)
        }
    "#,
    );
    assert_eq!(out, &["B"]);
}

#[test]
fn test_nested_local_class() {
    let out = run_prints(
        r#"
        fun main() {
            class Outer {
                class Inner(val v: Int)
            }
            println(Outer.Inner(1).v)
        }
    "#,
    );
    assert_eq!(out, &["1"]);
}

#[test]
fn test_nested_local_inner_class() {
    let out = run_prints(
        r#"
        fun main() {
            class Outer {
                inner class Inner(val base: String)
            }
            val o = Outer()
            println(o.Inner("x").base)
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_local_class_with_extension_fn() {
    let out = run_prints(
        r#"
        fun main() {
            class Word(val value: String)
            fun Word.quoted() = "\"$value\""
            println(Word("k").quoted())
        }
    "#,
    );
    assert_eq!(out, &["\"k\""]);
}

#[test]
fn test_local_class_object_expression() {
    let out = run_prints(
        r#"
        fun main() {
            val o = object {
                val v = 4
                fun text() = "v${v}"
            }
            println(o.text())
        }
    "#,
    );
    assert_eq!(out, &["v4"]);
}

#[test]
fn test_local_class_capture_outside_var() {
    let out = run_prints(
        r#"
        fun main() {
            val base = 10
            class Local {
                fun total(offset: Int) = base + offset
            }
            println(Local().total(3))
        }
    "#,
    );
    assert_eq!(out, &["13"]);
}

#[test]
fn test_local_recursive_class_method() {
    let out = run_prints(
        r#"
        fun main() {
            class Counter(val v: Int) {
                fun next(): Int = if (v <= 0) 0 else Counter(v - 1).next() + 1
            }
            println(Counter(3).next())
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_local_class_multiple_instances() {
    let out = run_prints(
        r#"
        fun main() {
            class Local {
                val x = 1
            }
            val a = Local()
            val b = Local()
            println(a.x + b.x)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_local_class_with_private() {
    let out = run_prints(
        r#"
        fun main() {
            class Box {
                private val hidden = 5
                fun reveal() = hidden
            }
            println(Box().reveal())
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_interface_and_impl() {
    let out = run_prints(
        r#"
        fun main() {
            interface I { fun value(): Int }
            class C(val v: Int) : I { override fun value() = v }
            println(C(4).value())
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_local_abstract_class() {
    let out = run_prints(
        r#"
        fun main() {
            abstract class A {
                abstract fun v(): Int
            }
            class B : A() {
                override fun v() = 6
            }
            println(B().v())
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_local_class_in_if_block() {
    let out = run_prints(
        r#"
        fun main() {
            val x = true
            if (x) {
                class Local(val v: Int)
                println(Local(7).v)
            }
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_local_class_in_when_block() {
    let out = run_prints(
        r#"
        fun main() {
            val n = 1
            val out = when (n) {
                1 -> {
                    class L(val v: String)
                    L("a").v
                }
                else -> "z"
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["a"]);
}

#[test]
fn test_local_class_with_secondary_constructor() {
    let out = run_prints(
        r#"
        fun main() {
            class Local {
                val v: Int
                constructor(v: Int) { this.v = v }
            }
            println(Local(5).v)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_class_with_init_block() {
    let out = run_prints(
        r#"
        fun main() {
            class Local {
                val v: Int
                init { v = 3 }
            }
            println(Local().v)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_local_class_type_alias_inside() {
    let out = run_prints(
        r#"
        fun main() {
            typealias Text = String
            class Local(val v: Text)
            println(Local("x").v)
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_local_class_array_like() {
    let out = run_prints(
        r#"
        fun main() {
            class Ints {
                val items = intArrayOf(1, 2, 3)
            }
            println(Ints().items.sum())
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_local_class_operator_fun() {
    let out = run_prints(
        r#"
        fun main() {
            class Box(val v: Int) {
                operator fun plus(other: Box) = Box(v + other.v)
            }
            println((Box(2) + Box(3)).v)
        }
    "#,
    );
    assert_eq!(out, &["5"]);
}

#[test]
fn test_local_class_and_local_function_interaction() {
    let out = run_prints(
        r#"
        fun main() {
            class C(val v: Int)
            fun f(v: C): Int = v.v * 2
            println(f(C(4)))
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_local_class_private_constructor_not_exposed() {
    let out = run_prints(
        r#"
        fun main() {
            class C private constructor(val v: Int) {
                companion object {
                    fun make(v: Int) = C(v)
                }
            }
            println(C.make(2).v)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_local_class_in_loop() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            for (i in 1..3) {
                class Local(val v: Int)
                out.append(Local(i).v)
            }
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["123"]);
}

#[test]
fn test_local_class_generic_property() {
    let out = run_prints(
        r#"
        fun main() {
            class Box<T>(val v: T)
            println(Box("x").v)
            println(Box(1).v)
        }
    "#,
    );
    assert_eq!(out, &["x", "1"]);
}

#[test]
fn test_local_class_in_lambda() {
    let out = run_prints(
        r#"
        fun main() {
            val v = { value: Int ->
                class Local(val doubled: Int)
                Local(value * 2)
            }
            println(v(3).doubled)
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_local_class_with_to_string_override() {
    let out = run_prints(
        r#"
        fun main() {
            class L(val v: Int) {
                override fun toString() = "val=" + v
            }
            println(L(8).toString())
        }
    "#,
    );
    assert_eq!(out, &["val=8"]);
}

#[test]
fn test_local_class_operator_plus() {
    let out = run_prints(
        r#"
        fun main() {
            class L(val v: Int) {
                operator fun inc() = L(v + 1)
            }
            println((L(1).inc().v))
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_local_class_with_nullable_field() {
    let out = run_prints(
        r#"
        fun main() {
            class L(val v: String?)
            println(L(null).v)
            println(L("x").v)
        }
    "#,
    );
    assert_eq!(out, &["null", "x"]);
}

#[test]
fn test_local_class_boolean_logic() {
    let out = run_prints(
        r#"
        fun main() {
            class Gate(val open: Boolean)
            println(Gate(true).open)
            println(Gate(false).open)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_local_class_list_composition() {
    let out = run_prints(
        r#"
        fun main() {
            class Node(val value: Int)
            val nodes = listOf(Node(1), Node(2))
            println(nodes[1].value)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}
