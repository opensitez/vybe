use crate::helpers::run_prints;

#[test]
fn test_method_dispatch_chooses_most_specific_override() {
    let out = run_prints(r#"
        open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
        }

        fun main() {
            val value: Base = Child()
            println(value.label())
        }
    "#);
    assert_eq!(out, &["child"]);
}

#[test]
fn test_field_access_uses_declared_reference_type() {
    let out = run_prints(r#"
        open class Base {
            val value: String = "base"
        }

        class Child : Base() {
            override val value: String = "child"
        }

        fun main() {
            val value = Child()
            val base: Base = value
            println(base.value)
            println(value.value)
        }
    "#);
    assert_eq!(out, &["child", "child"]);
}

#[test]
fn test_super_calls_parent_implementation() {
    let out = run_prints(r#"
        open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = super.label() + ":child"
        }

        fun main() {
            println(Child().label())
        }
    "#);
    assert_eq!(out, &["base:child"]);
}

#[test]
fn test_interface_dispatch_is_polymorphic() {
    let out = run_prints(r#"
        interface Reader {
            fun read(): String
        }

        class A : Reader {
            override fun read(): String = "a"
        }

        class B : Reader {
            override fun read(): String = "b"
        }

        fun emit(readers: Array<Reader>): String {
            var total = ""
            for (reader in readers) {
                total += reader.read()
            }
            return total
        }

        fun main() {
            println(emit(arrayOf(A(), B())))
        }
    "#);
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_multiple_interface_implementations_can_override_both() {
    let out = run_prints(r#"
        interface A {
            fun tag(): String = "A"
        }

        interface B {
            fun tag(): String = "B"
        }

        class C : A, B {
            override fun tag(): String = super<A>.tag() + "+" + super<B>.tag()
        }

        fun main() {
            println(C().tag())
        }
    "#);
    assert_eq!(out, &["A+B"]);
}

#[test]
fn test_abstract_dispatch_from_chain() {
    let out = run_prints(r#"
        abstract class Base {
            abstract fun emit(): Int
            open fun value(): Int = emit() * 2
        }

        class Child : Base() {
            override fun emit(): Int = 3
        }

        fun main() {
            val item: Base = Child()
            println(item.value())
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_open_property_can_be_mutated_in_child_override() {
    let out = run_prints(r#"
        open class Base {
            open var value: Int = 0
        }

        class Child : Base() {
            override var value: Int = 1
        }

        fun main() {
            val item = Child()
            item.value += 3
            println(item.value)
        }
    "#);
    assert_eq!(out, &["4"]);
}

#[test]
fn test_constructor_chain_preserves_virtual_dispatch() {
    let out = run_prints(r#"
        open class Base(value: Int) {
            init {
                if (value < 0) {
                    println("bad")
                }
            }
        }

        class Child(value: Int) : Base(value) {
            init {
                println("child")
            }
        }

        fun main() {
            Child(3)
        }
    "#);
    assert_eq!(out, &["child"]);
}

#[test]
fn test_subclass_without_override_uses_base_behavior() {
    let out = run_prints(r#"
        open class Base {
            open fun text(): String = "base"
        }

        class Direct : Base()

        fun main() {
            println(Direct().text())
        }
    "#);
    assert_eq!(out, &["base"]);
}

#[test]
fn test_generic_inheritance_dispatch_on_bounds() {
    let out = run_prints(r#"
        interface ValueCarrier {
            fun value(): Int
        }

        open class Base<T : ValueCarrier> : ValueCarrier {
            override fun value(): Int = 0
        }

        class Child : Base<Node>() {
            override fun value(): Int = 7
        }

        class Node : ValueCarrier {
            override fun value(): Int = 2
        }

        fun main() {
            val item: Base<*> = Child()
            val direct: ValueCarrier = Child()
            println(item.value())
            println(direct.value())
        }
    "#);
    assert_eq!(out, &["7", "7"]);
}
