use crate::helpers::run_prints;

#[test]
fn test_simple_class_delegation_forwards_call() {
    let out = run_prints(
        r#"
        interface Greeter { fun hello(): String }

        class Base(private val name: String) : Greeter {
            override fun hello() = "hello:$name"
        }

        class Wrapper(delegate: Greeter) : Greeter by delegate

        fun main() {
            val w = Wrapper(Base("kotlin"))
            println(w.hello())
        }
    "#,
    );
    assert_eq!(out, &["hello:kotlin"]);
}

#[test]
fn test_class_delegation_with_custom_members() {
    let out = run_prints(
        r#"
        interface Counter {
            fun value(): Int
        }

        class NumberCounter(private val v: Int) : Counter {
            override fun value() = v
        }

        class OffsetCounter(base: Counter) : Counter by base {
            fun id() = "offset"
        }

        fun main() {
            val c = OffsetCounter(NumberCounter(3))
            println(c.id())
            println(c.value())
        }
    "#,
    );
    assert_eq!(out, &["offset", "3"]);
}

#[test]
fn test_class_delegation_override_takes_precedence() {
    let out = run_prints(
        r#"
        interface Service { fun name(): String }

        class Primary : Service {
            override fun name() = "primary"
        }

        class Decorated(delegate: Service) : Service by delegate {
            override fun name() = "decorated"
        }

        fun main() {
            println(Decorated(Primary()).name())
        }
    "#,
    );
    assert_eq!(out, &["decorated"]);
}

#[test]
fn test_class_delegation_multiple_interface_members_forwarding() {
    let out = run_prints(
        r#"
        interface A { fun a(): String }
        interface B { fun b(): String }

        class Impl : A, B {
            override fun a() = "A"
            override fun b() = "B"
        }

        class Wrapper(private val impl: Impl) : A by impl, B by impl

        fun main() {
            val w = Wrapper(Impl())
            println(w.a())
            println(w.b())
        }
    "#,
    );
    assert_eq!(out, &["A", "B"]);
}

#[test]
fn test_list_interface_delegation_size_and_index() {
    let out = run_prints(
        r#"
        class ReadOnlyList(delegate: List<Int>) : List<Int> by delegate

        fun main() {
            val l = ReadOnlyList(listOf(2, 4, 6))
            println(l.size)
            println(l[1])
            println(l.contains(6))
        }
    "#,
    );
    assert_eq!(out, &["3", "4", "true"]);
}

#[test]
fn test_iterator_forwarded_from_list_delegation() {
    let out = run_prints(
        r#"
        class ReadOnlyList(delegate: List<String>) : Iterable<String> by delegate

        fun main() {
            val it: String = ReadOnlyList(listOf("a", "b")).joinToString("")
            println(it)
        }
    "#,
    );
    assert_eq!(out, &["ab"]);
}

#[test]
fn test_set_delegation_uses_set_contract() {
    let out = run_prints(
        r#"
        class ReadOnlySet(delegate: Set<Int>) : Set<Int> by delegate

        fun main() {
            val s = ReadOnlySet(setOf(1, 2, 2, 3))
            println(s.size)
            println(s.contains(2))
            println(s.contains(9))
        }
    "#,
    );
    assert_eq!(out, &["3", "true", "false"]);
}

#[test]
fn test_map_delegation_key_lookup() {
    let out = run_prints(
        r#"
        class ReadOnlyMap(delegate: Map<String, Int>) : Map<String, Int> by delegate

        fun main() {
            val m = ReadOnlyMap(mapOf("a" to 1, "b" to 2))
            println(m["a"])
            println(m.keys.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1", "a,b"]);
}

#[test]
fn test_delegation_with_base_property_access() {
    let out = run_prints(
        r#"
        interface Counter { val value: Int }

        class Base(override val value: Int) : Counter

        class Box(delegate: Counter) : Counter by delegate

        fun main() {
            println(Box(Base(7)).value)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_delegation_preserves_immutability_of_base_reference() {
    let out = run_prints(
        r#"
        interface View { fun size(): Int }

        class Snapshot(private val items: List<Int>) : View {
            override fun size() = items.size
        }

        class SnapshotWrapper(delegate: View) : View by delegate

        fun main() {
            val original = Snapshot(listOf(1, 2))
            val wrapped = SnapshotWrapper(original)
            println(wrapped.size())
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_class_delegation_with_varargs() {
    let out = run_prints(
        r#"
        interface Summer {
            fun sum(values: IntArray): Int
        }

        class Adder : Summer {
            override fun sum(values: IntArray): Int = values.sum()
        }

        class SumWrapper(delegate: Summer) : Summer by delegate

        fun main() {
            val wrapper = SumWrapper(Adder())
            println(wrapper.sum(intArrayOf(1, 2, 3)))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_collection_delegation_with_generic_type() {
    let out = run_prints(
        r#"
        interface Labeler<T> { fun label(value: T): String }

        class StringLabeler : Labeler<String> {
            override fun label(value: String) = "[$value]"
        }

        class DelegatingLabeler(delegate: Labeler<String>) : Labeler<String> by delegate

        fun main() {
            val l = DelegatingLabeler(StringLabeler())
            println(l.label("x"))
        }
    "#,
    );
    assert_eq!(out, &["[x]"]);
}

#[test]
fn test_nested_delegation_layer() {
    let out = run_prints(
        r#"
        interface Printer { fun print(value: Int): String }

        class BasePrinter : Printer {
            override fun print(value: Int): String = "base=$value"
        }

        class PrefixPrinter(delegate: Printer) : Printer by delegate
        class WrapperPrinter(delegate: Printer) : Printer by PrefixPrinter(delegate)

        fun main() {
            println(WrapperPrinter(BasePrinter()).print(4))
        }
    "#,
    );
    assert_eq!(out, &["base=4"]);
}

#[test]
fn test_delegate_object_expression() {
    let out = run_prints(
        r#"
        interface Messenger { fun message(): String }

        class Proxy(delegate: Messenger) : Messenger by delegate

        fun main() {
            val p = Proxy(object : Messenger {
                override fun message() = "from object"
            })
            println(p.message())
        }
    "#,
    );
    assert_eq!(out, &["from object"]);
}

#[test]
fn test_delegate_multiple_wrapped_calls_chain() {
    let out = run_prints(
        r#"
        interface Op { fun run(x: Int): Int }

        class A : Op { override fun run(x: Int) = x + 1 }
        class B(delegate: Op) : Op by delegate
        class C(delegate: Op) : Op by delegate

        fun main() {
            val c = C(B(A()))
            println(c.run(5))
        }
    "#,
    );
    assert_eq!(out, &["6"]);
}

#[test]
fn test_delegation_with_default_to_string_from_delegate() {
    let out = run_prints(
        r#"
        interface Named { fun title(): String }

        class NamedImpl : Named {
            override fun title() = "named"
            override fun toString() = "impl"
        }

        class NamedProxy(delegate: Named) : Named by delegate

        fun main() {
            val value = NamedProxy(NamedImpl())
            println(value.title())
            println(value.toString())
        }
    "#,
    );
    assert_eq!(out, &["named", "impl"]);
}

#[test]
fn test_delegation_with_nullable_delegate_reference() {
    let out = run_prints(
        r#"
        interface Marker { fun tag(): String }

        class MarkerImpl : Marker {
            override fun tag() = "ok"
        }

        class Holder(delegate: Marker) : Marker by delegate

        fun main() {
            val h = Holder(MarkerImpl())
            println(h.tag())
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_delegation_property_getter_forwarding() {
    let out = run_prints(
        r#"
        interface Counter {
            val count: Int
        }

        class State(override val count: Int) : Counter

        class Holder(delegate: Counter) : Counter by delegate

        fun main() {
            val value = Holder(State(12))
            println(value.count)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_delegation_with_custom_method_using_delegate() {
    let out = run_prints(
        r#"
        interface Adder { fun add(a: Int): Int }

        class BaseAdder : Adder {
            override fun add(a: Int) = a + 10
        }

        class WrapperAdder(delegate: Adder) : Adder by delegate {
            fun addTwice(a: Int): Int = add(a) + add(a)
        }

        fun main() {
            val value = WrapperAdder(BaseAdder())
            println(value.addTwice(4))
        }
    "#,
    );
    assert_eq!(out, &["28"]);
}

#[test]
fn test_delegation_with_collection_methods() {
    let out = run_prints(
        r#"
        interface Store {
            fun put(key: String, value: Int)
            fun get(key: String): Int?
        }

        class MemoryStore : Store {
            private val data = mutableMapOf<String, Int>()
            override fun put(key: String, value: Int) { data[key] = value }
            override fun get(key: String): Int? = data[key]
        }

        class StoreProxy(delegate: Store) : Store by delegate

        fun main() {
            val store = StoreProxy(MemoryStore())
            store.put("a", 3)
            println(store.get("a"))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_delegate_class_for_multiple_calls() {
    let out = run_prints(
        r#"
        interface Calc { fun step(v: Int): Int }

        class Base : Calc {
            override fun step(v: Int) = v + 1
        }

        class Wrapper(private val base: Calc) : Calc by base

        fun main() {
            val c = Wrapper(Base())
            println(c.step(0))
            println(c.step(9))
        }
    "#,
    );
    assert_eq!(out, &["1", "10"]);
}

#[test]
fn test_delegate_when_base_state_changes() {
    let out = run_prints(
        r#"
        interface MutableCounter { var value: Int }

        class Counter(var value: Int) : MutableCounter

        class Proxy(delegate: MutableCounter) : MutableCounter by delegate

        fun main() {
            val c = Counter(1)
            val p = Proxy(c)
            p.value = p.value + 2
            println(p.value)
            println(c.value)
        }
    "#,
    );
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_delegate_chain_keeps_same_reference() {
    let out = run_prints(
        r#"
        interface Marker { fun marker(): String }

        class First : Marker { override fun marker() = "first" }
        class Second(delegate: Marker) : Marker by delegate
        class Third(delegate: Marker) : Marker by delegate

        fun main() {
            val t = Third(Second(First()))
            println(t.marker())
        }
    "#,
    );
    assert_eq!(out, &["first"]);
}

#[test]
fn test_delegate_with_extensionless_interface() {
    let out = run_prints(
        r#"
        interface Identity { fun id(): String }

        class A : Identity { override fun id() = "A" }
        class Holder(delegate: Identity) : Identity by delegate

        fun main() {
            val h = Holder(A())
            println(h.id())
        }
    "#,
    );
    assert_eq!(out, &["A"]);
}

#[test]
fn test_delegate_inheritance_not_allowed_not_used() {
    let out = run_prints(
        r#"
        interface Base { fun kind(): String }

        open class Root : Base {
            override fun kind() = "root"
        }

        class Child(delegate: Base) : Base by delegate, Root()

        fun main() {
            println(Child(Root()).kind())
        }
    "#,
    );
    assert_eq!(out, &["root"]);
}

#[test]
fn test_delegate_with_custom_extension_function_usage() {
    let out = run_prints(
        r#"
        interface Text { fun text(): String }

        class Source : Text {
            override fun text() = "value"
        }

        class Delegate(delegate: Text) : Text by delegate

        fun Text.enhancedSuffix(): String = text() + "!"

        fun main() {
            val d = Delegate(Source())
            println(d.enhancedSuffix())
        }
    "#,
    );
    assert_eq!(out, &["value!"]);
}
