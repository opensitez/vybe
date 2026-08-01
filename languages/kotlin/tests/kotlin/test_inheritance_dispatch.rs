use crate::helpers::run_prints;

#[test]
fn test_method_dispatch_chooses_most_specific_override() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["child"]);
}

#[test]
fn test_field_access_uses_declared_reference_type() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["child", "child"]);
}

#[test]
fn test_super_calls_parent_implementation() {
    let out = run_prints(
        r#"
        open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = super.label() + ":child"
        }

        fun main() {
            println(Child().label())
        }
    "#,
    );
    assert_eq!(out, &["base:child"]);
}

#[test]
fn test_interface_dispatch_is_polymorphic() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_multiple_interface_implementations_can_override_both() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["A+B"]);
}

#[test]
fn test_abstract_dispatch_from_chain() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_open_property_can_be_mutated_in_child_override() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_constructor_chain_preserves_virtual_dispatch() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["child"]);
}

#[test]
fn test_subclass_without_override_uses_base_behavior() {
    let out = run_prints(
        r#"
        open class Base {
            open fun text(): String = "base"
        }

        class Direct : Base()

        fun main() {
            println(Direct().text())
        }
    "#,
    );
    assert_eq!(out, &["base"]);
}

#[test]
fn test_generic_inheritance_dispatch_on_bounds() {
    let out = run_prints(
        r#"
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
    "#,
    );
    assert_eq!(out, &["7", "7"]);
}

#[test]
fn test_virtual_method_is_dispatched_from_parent_reference() {
    let out = run_prints(
        r#"
        open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
        }

        fun main() {
            val item: Base = Child()
            println(item.label())
        }
    "#,
    );
    assert_eq!(out, &["child"]);
}

#[test]
fn test_virtual_property_dispatch_in_inheritance_chain() {
    let out = run_prints(
        r#"
        open class Base {
            open val value: Int = 1
            open fun total(): Int = value + 1
        }

        class Child : Base() {
            override val value: Int = 3
            override fun total(): Int = value + 2
        }

        fun main() {
            val item: Base = Child()
            println(item.value)
            println(item.total())
        }
    "#,
    );
    assert_eq!(out, &["3", "5"]);
}

#[test]
fn test_interface_default_implementation_can_be_overridden() {
    let out = run_prints(
        r#"
        interface Messenger {
            fun text(): String = "default"
        }

        class Custom : Messenger {
            override fun text(): String = "custom"
        }

        class InheritDefault : Messenger

        fun main() {
            println(Custom().text())
            println(InheritDefault().text())
        }
    "#,
    );
    assert_eq!(out, &["custom", "default"]);
}

#[test]
fn test_interface_conflict_resolution_with_two_defaults() {
    let out = run_prints(
        r#"
        interface A {
            fun text(): String = "a"
        }

        interface B {
            fun text(): String = "b"
        }

        class C : A, B {
            override fun text(): String = super<A>.text() + "," + super<B>.text()
        }

        fun main() {
            println(C().text())
        }
    "#,
    );
    assert_eq!(out, &["a,b"]);
}

#[test]
fn test_abstract_overrides_are_required_by_concrete_class() {
    let out = run_prints(
        r#"
        abstract class Base {
            abstract val title: String
        }

        class Leaf : Base() {
            override val title: String = "leaf"
        }

        fun main() {
            println(Leaf().title)
        }
    "#,
    );
    assert_eq!(out, &["leaf"]);
}

#[test]
fn test_super_call_in_overridden_method_uses_parent_behavior() {
    let out = run_prints(
        r#"
        open class Base {
            open fun score(value: Int): Int = value * 2
        }

        class Child : Base() {
            override fun score(value: Int): Int = super.score(value) + 1
        }

        fun main() {
            println(Child().score(3))
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_multiple_class_levels_share_virtual_method() {
    let out = run_prints(
        r#"
        open class Base {
            open fun route(): String = "base"
        }

        open class Mid : Base() {
            override fun route(): String = "mid"
        }

        class Leaf : Mid() {
            override fun route(): String = "leaf"
        }

        fun emit(route: Base): String = route.route()

        fun main() {
            println(emit(Base()))
            println(emit(Mid()))
            println(emit(Leaf()))
        }
    "#,
    );
    assert_eq!(out, &["base", "mid", "leaf"]);
}

#[test]
fn test_base_class_members_remain_when_override() {
    let out = run_prints(
        r#"
        open class Base {
            open fun label(): String = "base"
        }

        class Child : Base() {
            override fun label(): String = "child"
            fun asBase(): Base = this
        }

        fun main() {
            val child = Child()
            println(child.label())
            println(child.asBase().label())
        }
    "#,
    );
    assert_eq!(out, &["child", "child"]);
}

#[test]
fn test_open_var_override_preserves_mutability_contract() {
    let out = run_prints(
        r#"
        open class Base {
            open var value: Int = 1
        }

        class Child : Base() {
            override var value: Int = 5
        }

        fun main() {
            val child = Child()
            child.value += 2
            println(child.value)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_implementation_of_multiple_interfaces_is_dispatched_as_expected() {
    let out = run_prints(
        r#"
        interface Read {
            fun read(): String = "read"
        }

        interface Write {
            fun write(): String = "write"
        }

        class Device : Read, Write {
            override fun read(): String = "device-read"
            override fun write(): String = "device-write"
        }

        fun main() {
            val device: Read = Device()
            val writer: Write = Device()
            println(device.read())
            println(writer.write())
        }
    "#,
    );
    assert_eq!(out, &["device-read", "device-write"]);
}

#[test]
fn test_generic_dispatch_on_inheritance() {
    let out = run_prints(
        r#"
        interface Box {
            fun value(): Int
        }

        open class Holder<T : Box> : Box {
            override fun value(): Int = 0
        }

        class Fast : Holder<IntBox>() {
            override fun value(): Int = 4
        }

        class IntBox : Box {
            override fun value(): Int = 9
        }

        fun main() {
            val item: Holder<*> = Fast()
            val typed: Box = Fast()
            println(item.value())
            println(typed.value())
        }
    "#,
    );
    assert_eq!(out, &["4", "4"]);
}

#[test]
fn test_polymorphic_array_dispatch_processes_dynamic_types() {
    let out = run_prints(
        r#"
        open class Node {
            open fun kind(): String = "node"
        }

        class Leaf : Node() {
            override fun kind(): String = "leaf"
        }

        class Branch : Node() {
            override fun kind(): String = "branch"
        }

        fun summarize(nodes: Array<Node>): String {
            var value = ""
            for (node in nodes) {
                value += node.kind()
                value += ";"
            }
            return value
        }

        fun main() {
            val nodes: Array<Node> = arrayOf(Node(), Leaf(), Branch())
            println(summarize(nodes))
        }
    "#,
    );
    assert_eq!(out, &["node;leaf;branch;"]);
}

#[test]
fn test_dispatch_from_rebound_base_reference() {
    let out = run_prints(
        r#"
        open class Printer {
            open fun emit(prefix: String): String = prefix + ":base"
        }

        class Loud : Printer() {
            override fun emit(prefix: String): String = prefix + ":loud"
        }

        fun emitTwice(printer: Printer): String {
            return printer.emit("x") + "," + printer.emit("y")
        }

        fun main() {
            var printer: Printer = Printer()
            printer = Loud()
            println(emitTwice(printer))
        }
    "#,
    );
    assert_eq!(out, &["x:loud,y:loud"]);
}

#[test]
fn test_override_chain_across_multiple_levels() {
    let out = run_prints(
        r#"
        open class Base {
            open fun route(): String = "base"
        }

        open class Mid : Base() {
            override fun route(): String = "mid" + super.route()
        }

        class Leaf : Mid() {
            override fun route(): String = "leaf" + super.route()
        }

        fun main() {
            val node: Base = Leaf()
            println(node.route())
        }
    "#,
    );
    assert_eq!(out, &["leafmidbase"]);
}

#[test]
fn test_casting_between_class_and_interface_preserves_dispatch() {
    let out = run_prints(
        r#"
        interface Labeled {
            fun label(): String = "interface"
        }

        open class Base {
            open fun label(): String = "base"
        }

        class Widget : Base(), Labeled {
            override fun label(): String = "widget"
        }

        fun callThroughInterface(value: Labeled): String {
            return value.label()
        }

        fun main() {
            val root: Base = Widget()
            val viaInterface = root as Labeled
            println((root as Widget).label())
            println(callThroughInterface(viaInterface))
        }
    "#,
    );
    assert_eq!(out, &["widget", "widget"]);
}

#[test]
fn test_getter_and_setter_overrides_are_used_from_base_reference() {
    let out = run_prints(
        r#"
        open class Base {
            open var value: Int = 0
        }

        class Child : Base() {
            private var storage = 10

            override var value: Int
                get() = storage
                set(new_value) {
                    storage = new_value + 1
                }
        }

        fun main() {
            val base: Base = Child()
            base.value = 7
            println(base.value)
            println((base as Child).value)
        }
    "#,
    );
    assert_eq!(out, &["8", "8"]);
}

#[test]
fn test_interface_default_method_can_be_selected_with_super() {
    let out = run_prints(
        r#"
        interface Counter {
            fun value(): String = "default"
        }

        class Custom : Counter {
            override fun value(): String = super<Counter>.value() + "-custom"
        }

        class Plain : Counter

        fun main() {
            val custom: Counter = Custom()
            val plain: Counter = Plain()
            println(custom.value())
            println(plain.value())
        }
    "#,
    );
    assert_eq!(out, &["default-custom", "default"]);
}

#[test]
fn test_dispatch_in_higher_order_function_argument() {
    let out = run_prints(
        r#"
        open class Labelled {
            open fun text(): String = "base"
        }

        class Dynamic : Labelled() {
            override fun text(): String = "dynamic"
        }

        fun mapLabel(value: Labelled, render: (Labelled) -> String): String {
            return render(value)
        }

        fun main() {
            val item: Labelled = Dynamic()
            println(mapLabel(item) { it.text() })
            println(mapLabel(item) { target -> "[" + target.text() + "]" })
        }
    "#,
    );
    assert_eq!(out, &["dynamic", "[dynamic]"]);
}

#[test]
fn test_interface_default_with_class_super_resolution() {
    let out = run_prints(
        r#"
        interface Tracer {
            fun route(): String = "trace"
        }

        open class Base {
            open fun route(): String = "base"
        }

        class Logger : Base(), Tracer {
            override fun route(): String = super<Tracer>.route() + ":" + super<Base>.route()
        }

        fun main() {
            val logger: Base = Logger()
            println(logger.route())
            println((logger as Logger).route())
        }
    "#,
    );
    assert_eq!(out, &["trace:base", "trace:base"]);
}

#[test]
fn test_generic_override_dispatches_by_dynamic_type() {
    let out = run_prints(
        r#"
        open class Base<T> {
            open fun describe(value: T): String = "base:" + value.toString()
        }

        class Text : Base<String>() {
            override fun describe(value: String): String = "text:" + value
        }

        fun main() {
            val item: Base<String> = Text()
            println(item.describe("ok"))
            println((item as Text).describe("x"))
        }
    "#,
    );
    assert_eq!(out, &["text:ok", "text:x"]);
}

#[test]
fn test_chained_super_calls_build_expected_behavior() {
    let out = run_prints(
        r#"
        open class Base {
            open fun label(): String = "base"
        }

        open class Mid : Base() {
            override fun label(): String = super.label() + ":mid"
        }

        class Leaf : Mid() {
            override fun label(): String = super.label() + ":leaf"
        }

        fun main() {
            val value: Base = Leaf()
            println(value.label())
        }
    "#,
    );
    assert_eq!(out, &["base:mid:leaf"]);
}
