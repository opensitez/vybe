use crate::helpers::run_prints;

#[test]
fn test_variance_covariant_producer_string_to_any() {
    let out = run_prints(r#"
        interface Producer<out T> { fun provide(): T }
        class StringSource : Producer<String> {
            override fun provide(): String = "x"
        }
        fun main() {
            val source: Producer<Any> = StringSource()
            println(source.provide())
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_variance_covariant_read_only_list_to_iterable() {
    let out = run_prints(r#"
        val values: List<String> = listOf("a", "b")
        val anyValues: List<Any> = values
        println(anyValues.size)
    "#);
    assert_eq!(out, &["2"]);
}

#[test]
fn test_variance_invariant_blocked_assignment() {
    let out = run_prints(r#"
        open class Animal
        class Dog : Animal()
        fun <T> accepts(list: MutableList<T>) {
            println(list.size)
        }
        fun main() {
            val dogs: MutableList<Dog> = mutableListOf(Dog())
            accepts(dogs)
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_variance_contravariant_sink_animal() {
    let out = run_prints(r#"
        interface Sink<in T> { fun consume(value: T) }
        open class Animal { fun label() = "a" }
        class Recorder : Sink<Animal> {
            override fun consume(value: Animal) { println(value.label()) }
        }
        class Dog : Animal() { override fun label() = "d" }
        fun main() {
            val sink: Sink<Dog> = Recorder()
            sink.consume(Dog())
        }
    "#);
    assert_eq!(out, &["d"]);
}

#[test]
fn test_variance_variance_in_array_like_class() {
    let out = run_prints(r#"
        class Box<out T>(val value: T)
        fun main() {
            val textBox: Box<String> = Box("v")
            val item: Box<Any> = textBox
            println(item.value)
        }
    "#);
    assert_eq!(out, &["v"]);
}

#[test]
fn test_variance_inout_projection_get() {
    let out = run_prints(r#"
        fun firstOrNull(items: List<out Any>): String = items.firstOrNull()?.toString() ?: "none"
        fun main() {
            println(firstOrNull(listOf(1, 2)))
            println(firstOrNull(listOf("x")))
        }
    "#);
    assert_eq!(out, &["1", "x"]);
}

#[test]
fn test_variance_inout_projection_set_forbidden_path_not_available() {
    let out = run_prints(r#"
        fun accepts(outItems: List<out String>) {
            println(outItems.size)
        }
        fun main() {
            accepts(listOf("a", "b", "c"))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_variance_function_input_type_upcast() {
    let out = run_prints(r#"
        open class Animal
        class Cat : Animal()
        class Dog : Animal()
        fun feed(animal: Animal) { println(animal::class.simpleName) }
        fun main() {
            val cat: Cat = Cat()
            val dog: Dog = Dog()
            feed(cat)
            feed(dog)
        }
    "#);
    assert_eq!(out, &["Cat", "Dog"]);
}

#[test]
fn test_variance_generic_function_out_type() {
    let out = run_prints(r#"
        interface Source<out T> {
            fun get(): T
        }
        class UserSource : Source<String> {
            override fun get(): String = "u"
        }
        fun wrap(source: Source<out Any>): Any {
            return source.get()
        }
        fun main() {
            println(wrap(UserSource()))
        }
    "#);
    assert_eq!(out, &["u"]);
}

#[test]
fn test_variance_generic_function_in_type() {
    let out = run_prints(r#"
        interface Writer<in T> {
            fun write(value: T)
        }
        open class Thing
        class ThingWriter : Writer<Thing> {
            override fun write(value: Thing) { println("w") }
        }
        fun consumeWriter(writer: Writer<Thing>) {
            writer.write(Thing())
        }
        class Fancy : Thing()
        fun main() {
            val writer: Writer<Fancy> = ThingWriter()
            consumeWriter(writer)
        }
    "#);
    assert_eq!(out, &["w"]);
}

#[test]
fn test_variance_list_projection_safe_read() {
    let out = run_prints(r#"
        fun readFirst(items: List<out Number>): Int {
            return items.first().toInt()
        }
        fun main() {
            println(readFirst(listOf<Int>(5, 7)))
            println(readFirst(listOf<Long>(9L, 10L)))
        }
    "#);
    assert_eq!(out, &["5", "9"]);
}

#[test]
fn test_variance_nested_generics_covariant_nested() {
    let out = run_prints(r#"
        interface Boxed<out T> { val value: T }
        class Holder<T>(override val value: T) : Boxed<T>
        fun main() {
            val boxed: Boxed<String> = Holder("k")
            val anyBox: Boxed<Any> = boxed
            println(anyBox.value)
        }
    "#);
    assert_eq!(out, &["k"]);
}

#[test]
fn test_variance_nested_generics_contravariant_nested() {
    let out = run_prints(r#"
        interface Sink<in T> { fun consume(v: T) }
        open class Node
        class Leaf : Node()
        class NodeSink : Sink<Node> { override fun consume(v: Node) { println(v::class.simpleName) } }
        fun main() {
            val sink: Sink<Leaf> = NodeSink()
            sink.consume(Leaf())
        }
    "#);
    assert_eq!(out, &["Leaf"]);
}

#[test]
fn test_variance_star_projection_read() {
    let out = run_prints(r#"
        fun read(items: List<*>, idx: Int): String {
            return items[idx]?.toString() ?: "nil"
        }
        fun main() {
            println(read(listOf("x", "y"), 0))
            println(read(listOf<Int>(1, 2), 1))
        }
    "#);
    assert_eq!(out, &["x", "2"]);
}

#[test]
fn test_variance_star_projection_write_not_attempted() {
    let out = run_prints(r#"
        fun count(items: MutableList<*>) : Int = items.size
        fun main() {
            println(count(mutableListOf(1, 2, 3)))
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_variance_projection_with_transform() {
    let out = run_prints(r#"
        fun stringify(values: List<out Any?>): String = values.joinToString(",")
        fun main() {
            println(stringify(listOf(1, "a", true)))
        }
    "#);
    assert_eq!(out, &["1,a,true"]);
}

#[test]
fn test_variance_invariant_list_copy() {
    let out = run_prints(r#"
        fun copyAll(src: List<out Number>, dst: MutableList<in Number>) {
            src.forEach { dst.add(it.toInt()) }
            println(dst)
        }
        fun main() {
            val output = mutableListOf<Number>()
            copyAll(listOf(1, 2, 3), output)
        }
    "#);
    assert_eq!(out, &["[1, 2, 3]"]);
}

#[test]
fn test_variance_function_return_type_covariant() {
    let out = run_prints(r#"
        open class Base
        class Child : Base()
        fun make(): Child = Child()
        fun produce(): Base = make()
        fun main() {
            val b: Base = produce()
            println(b is Base)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_variance_type_projection_on_map_key() {
    let out = run_prints(r#"
        fun keyCount(map: Map<out String, *>) = map.size
        fun main() {
            println(keyCount(mapOf("a" to 1, "b" to 2)))
            println(keyCount(mapOf("z" to "x")))
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_variance_map_variance_values_projection() {
    let out = run_prints(r#"
        fun collectValues(values: Map<*, out Number>): Int {
            return values.values.sumBy { it.toInt() }
        }
        fun main() {
            println(collectValues(mapOf("a" to 1, "b" to 2)))
            println(collectValues(mapOf("x" to 3L)))
        }
    "#);
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_variance_out_projection_setter_not_allowed_compile_time_skipped() {
    let out = run_prints(r#"
        class Repo<T> {
            private val store = mutableListOf<T>()
            fun getStore(): List<T> = store
            fun addAll(values: List<out T>) {
                // this path is intentionally empty
            }
            fun mainAdd() {
                println(store.size)
            }
        }
        fun main() {
            val r = Repo<String>()
            r.mainAdd()
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_variance_consumer_receiver_accepts_subtype() {
    let out = run_prints(r#"
        interface Consume<in T> { fun accept(v: T) }
        open class Item
        class Special : Item()
        class Sink : Consume<Item> { override fun accept(v: Item) { println(v::class.simpleName) } }
        fun main() {
            val consume: Consume<Special> = Sink()
            consume.accept(Special())
        }
    "#);
    assert_eq!(out, &["Special"]);
}

#[test]
fn test_variance_producer_receiver_returns_subtype() {
    let out = run_prints(r#"
        interface Produce<out T> { fun next(): T }
        class SpecialProducer : Produce<Special> {
            override fun next(): Special = Special()
        }
        open class Special
        open class Item
        fun main() {
            val p: Produce<Item> = SpecialProducer()
            println(p.next() is Special)
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_variance_mutable_projection_read_write_split() {
    let out = run_prints(r#"
        fun copyValues(source: List<out Int>, target: MutableList<Int>) {
            source.forEach { target.add(it) }
            println(target.joinToString(","))
        }
        fun main() {
            val target = mutableListOf<Int>()
            copyValues(listOf(1, 2), target)
        }
    "#);
    assert_eq!(out, &["1,2"]);
}

#[test]
fn test_variance_map_kotlin_read_projection() {
    let out = run_prints(r#"
        fun readAny(map: Map<String, out Number>): String = map.values.joinToString("-") { it.toString() }
        fun main() {
            println(readAny(mapOf("a" to 1)))
            println(readAny(mapOf("b" to 2.0)))
        }
    "#);
    assert_eq!(out, &["1", "2.0"]);
}

#[test]
fn test_variance_nested_projection_pair_first() {
    let out = run_prints(r#"
        val pairFirst: (Pair<out String, *>) -> String = { p -> p.first }
        fun main() {
            println(pairFirst(Pair("x", 3)))
        }
    "#);
    assert_eq!(out, &["x"]);
}

#[test]
fn test_variance_nested_projection_pair_second_readonly() {
    let out = run_prints(r#"
        val pairSecond: (Pair<*, out Number>) -> String = { p -> p.second.toString() }
        fun main() {
            println(pairSecond(Pair("x", 9L)))
            println(pairSecond(Pair(1, 4.5)))
        }
    "#);
    assert_eq!(out, &["9", "4.5"]);
}

#[test]
fn test_variance_generics_with_projection_in_class() {
    let out = run_prints(r#"
        class Box<T>(val value: T)
        fun printBox(values: Box<out Any>) {
            println(values.value)
        }
        fun main() {
            printBox(Box("hello"))
            printBox(Box(99))
        }
    "#);
    assert_eq!(out, &["hello", "99"]);
}

#[test]
fn test_variance_generics_with_projection_in_function_params() {
    let out = run_prints(r#"
        fun pick(values: List<out Any>): String {
            return values[0].toString()
        }
        fun main() {
            println(pick(listOf(1, 2)))
            println(pick(listOf("x", "y")))
        }
    "#);
    assert_eq!(out, &["1", "x"]);
}

#[test]
fn test_variance_transform_out_to_in() {
    let out = run_prints(r#"
        interface Producer<out T> { fun get(): T }
        interface Consumer<in T> { fun put(v: T) }
        class StringProducer : Producer<String> {
            override fun get(): String = "ok"
        }
        class Printer : Consumer<Any> {
            override fun put(v: Any) { println(v.toString()) }
        }
        fun main() {
            val producer: Producer<Any> = StringProducer()
            val consumer: Consumer<String> = Printer()
            consumer.put(producer.get())
        }
    "#);
    assert_eq!(out, &["ok"]);
}
