kotlin_run_test!(
    test_nested_class_basic_access,
    r#"
        class Box {
            class Inner(val value: Int)
        }

        fun main() {
            val item = Box.Inner(7)
            println(item.value)
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_nested_class_generic_type,
    r#"
        class Holder<T> {
            class Slot<T>(val payload: T)
        }

        fun main() {
            val slot = Holder.Slot("text")
            println(slot.payload)
        }
    "#,
    &["text"]
);

kotlin_run_test!(
    test_nested_class_hierarchy,
    r#"
        class Graph {
            class Node(val label: String)
            class Edge(val from: String, val to: String)
        }

        fun main() {
            val node = Graph.Node("a")
            val edge = Graph.Edge("a", "b")
            println(node.label)
            println(edge.from + edge.to)
        }
    "#,
    &["a", "ab"]
);

kotlin_run_test!(
    test_nested_class_with_function_calls,
    r#"
        class Registry {
            class Entry(val name: String)

            companion object {
                fun make(name: String): Entry = Entry(name)
            }
        }

        fun main() {
            val e = Registry.make("x")
            println(e.name)
        }
    "#,
    &["x"]
);

kotlin_run_test!(
    test_nested_class_method_dispatch,
    r#"
        class Service {
            class State(val ok: Boolean)
        }

        fun main() {
            val a = Service.State(true)
            val b = Service.State(false)
            val total = (if (a.ok) 1 else 0) + (if (b.ok) 1 else 0)
            println(total)
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_nested_class_in_function_scope,
    r#"
        fun factory(): String {
            class Packet(val value: Int)
            return Packet(3).value.toString()
        }

        fun main() {
            println(factory())
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_nested_class_in_object,
    r#"
        object Bridge {
            class Board(val size: Int)
            fun board(): Board = Board(4)
        }

        fun main() {
            println(Bridge.board().size)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_nested_class_reference_chain,
    r#"
        class Factory {
            class Producer {
                class Widget(val label: String)
            }
        }

        fun main() {
            val widget = Factory.Producer.Widget("gear")
            println(widget.label)
        }
    "#,
    &["gear"]
);

kotlin_run_test!(
    test_nested_class_in_generic_container,
    r#"
        class Bag<T> {
            class Entry
            class Typed<T>(val value: T)
        }

        fun main() {
            val a = Bag.Typed(2)
            val b = Bag.Entry()
            println(a.value + 1)
            println(b::class.simpleName)
        }
    "#,
    &["3", "Entry"]
);

kotlin_run_test!(
    test_nested_class_local_type_identity,
    r#"
        class Container {
            class Marker
        }

        fun main() {
            val first: Container.Marker = Container.Marker()
            val second = Container.Marker()
            println(first == second)
            println(first != null)
        }
    "#,
    &["false", "true"]
);

kotlin_run_test!(
    test_nested_class_with_static_init_count,
    r#"
        class Counter {
            class Item {
                companion object { var count = 0 }
            }

            init {
                Item.count += 1
            }
        }

        fun main() {
            Counter()
            Counter()
            println(Counter.Item.count)
        }
    "#,
    &["2"]
);
