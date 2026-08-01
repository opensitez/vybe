use crate::helpers::run_prints;

#[test]
fn test_generic_identity_function() {
    let out = run_prints(
        r#"
        fun <T> identity(value: T): T {
            return value
        }

        fun main() {
            println(identity(3))
            println(identity("ok"))
        }
    "#,
    );
    assert_eq!(out, &["3", "ok"]);
}

#[test]
fn test_generic_pair_data_holder() {
    let out = run_prints(
        r#"
        class Holder<K, V>(private val key: K, private val value: V) {
            fun parts(): String {
                return key.toString() + ":" + value.toString()
            }
        }

        fun main() {
            println(Holder("x", 7).parts())
            println(Holder(2, true).parts())
        }
    "#,
    );
    assert_eq!(out, &["x:7", "2:true"]);
}

#[test]
fn test_generic_list_first_or_default() {
    let out = run_prints(
        r#"
        fun <T> firstOrDefault(values: Array<T>, fallback: T): T {
            if (values.size == 0) {
                return fallback
            }
            return values[0]
        }

        fun main() {
            println(firstOrDefault(arrayOf(4, 7, 9), 0))
            println(firstOrDefault(arrayOf<String>(), "none"))
        }
    "#,
    );
    assert_eq!(out, &["4", "none"]);
}

#[test]
fn test_generic_type_bound_number() {
    let out = run_prints(
        r#"
        fun <T : Number> asInt(value: T): Int {
            return value.toInt()
        }

        fun main() {
            println(asInt(12))
            println(asInt(12.7))
        }
    "#,
    );
    assert_eq!(out, &["12", "12"]);
}

#[test]
fn test_generic_multiple_type_parameters() {
    let out = run_prints(
        r#"
        fun <K, V> choose(left: K, right: V): String {
            return left.toString() + right.toString()
        }

        fun main() {
            println(choose("a", 1))
            println(choose(2, "b"))
        }
    "#,
    );
    assert_eq!(out, &["a1", "2b"]);
}

#[test]
fn test_generic_reified_like_inference_from_arguments() {
    let out = run_prints(
        r#"
        fun <T> asStringList(value: T, converter: (T) -> String): String {
            return converter(value)
        }

        fun main() {
            println(asStringList(3, { it.toString() }))
            println(asStringList(false, { v -> v.toString() }))
        }
    "#,
    );
    assert_eq!(out, &["3", "false"]);
}

#[test]
fn test_generic_constraint_comparable_max() {
    let out = run_prints(
        r#"
        fun <T : Comparable<T>> maxOf(first: T, second: T): T {
            return if (first > second) first else second
        }

        fun main() {
            println(maxOf(4, 9))
            println(maxOf("a", "z"))
        }
    "#,
    );
    assert_eq!(out, &["9", "z"]);
}

#[test]
fn test_generic_class_with_method() {
    let out = run_prints(
        r#"
        class Cache<T>(initial: T) {
            private val value: T = initial
            fun unwrap(): T {
                return value
            }
        }

        fun main() {
            println(Cache("hello").unwrap())
            println(Cache(8).unwrap())
        }
    "#,
    );
    assert_eq!(out, &["hello", "8"]);
}

#[test]
fn test_generic_extension_function() {
    let out = run_prints(
        r#"
        class Holder<T>(val value: T)

        fun <T> Holder<T>.labeled(prefix: String): String {
            return prefix + ":" + this.value.toString()
        }

        fun main() {
            println(Holder(5).labeled("x"))
            println(Holder("one").labeled("y"))
        }
    "#,
    );
    assert_eq!(out, &["x:5", "y:one"]);
}

#[test]
fn test_generic_array_projection_alias() {
    let out = run_prints(
        r#"
        fun <T> repeatThree(value: T): Array<T> {
            return arrayOf(value, value, value)
        }

        fun main() {
            val values = repeatThree("go")
            println(values[0])
            println(values[1])
            println(values[2])
        }
    "#,
    );
    assert_eq!(out, &["go", "go", "go"]);
}

#[test]
fn test_generic_factory_from_literal() {
    let out = run_prints(
        r#"
        fun <T> one(value: T): Array<T> {
            return arrayOf(value)
        }

        fun main() {
            val numbers = one(42)
            val words = one("zap")
            println(numbers.size)
            println(numbers[0])
            println(words[0])
        }
    "#,
    );
    assert_eq!(out, &["1", "42", "zap"]);
}

#[test]
fn test_generic_bound_charsequence_length() {
    let out = run_prints(
        r#"
        fun <T : CharSequence> totalLength(left: T, right: T): Int {
            return left.length + right.length
        }

        fun main() {
            println(totalLength("ab", "xyz"))
            println(totalLength(StringBuilder("k"), StringBuilder("on")))
        }
    "#,
    );
    assert_eq!(out, &["5", "3"]);
}

#[test]
fn test_generic_numeric_projection_to_double() {
    let out = run_prints(
        r#"
        fun <T : Number> sumToDouble(values: Array<T>): Double {
            var total = 0.0
            for (value in values) {
                total += value.toDouble()
            }
            return total
        }

        fun main() {
            val ints = arrayOf(1, 2, 3)
            val doubles = arrayOf(1.5, 2.5)
            println(sumToDouble(ints))
            println(sumToDouble(doubles))
        }
    "#,
    );
    assert_eq!(out, &["6", "4"]);
}

#[test]
fn test_generic_interface_contract() {
    let out = run_prints(
        r#"
        interface Provider<T> {
            fun get(): T
        }

        class Constant<T>(private val value: T) : Provider<T> {
            override fun get(): T = value
        }

        fun <T> read(provider: Provider<T>): T {
            return provider.get()
        }

        fun main() {
            val name = Constant("Alice")
            val number = Constant(77)
            println(read(name))
            println(read(number))
        }
    "#,
    );
    assert_eq!(out, &["Alice", "77"]);
}

#[test]
fn test_generic_extension_constraint_and_mapping() {
    let out = run_prints(
        r#"
        class Wrapper<T>(val value: T)

        fun <T> Wrapper<T>.map(transform: (T) -> T): T {
            return transform(this.value)
        }

        fun main() {
            println(Wrapper("a").map { it + "b" })
            println(Wrapper(9).map { it + 1 })
        }
    "#,
    );
    assert_eq!(out, &["ab", "10"]);
}

#[test]
fn test_generic_list_projection_readonly_access() {
    let out = run_prints(
        r#"
        fun <T> joinValues(values: List<out T>): String {
            var out = ""
            for (value in values) {
                out += value.toString()
            }
            return out
        }

        fun main() {
            val ints = listOf(1, 2, 3)
            val texts: List<String> = listOf("a", "b")
            println(joinValues(ints))
            println(joinValues(texts))
        }
    "#,
    );
    assert_eq!(out, &["123", "ab"]);
}

#[test]
fn test_generic_writable_projection() {
    let out = run_prints(
        r#"
        fun <T> appendDefault(values: MutableList<in T>, value: T) {
            values.add(value)
        }

        fun main() {
            val numbers: MutableList<Number> = mutableListOf(1)
            appendDefault<Number>(numbers, 2)
            appendDefault<Number>(numbers, 3.5)
            println(numbers[1])
            println(numbers[2])
        }
    "#,
    );
    assert_eq!(out, &["2", "3.5"]);
}

#[test]
fn test_generic_out_variance_reader() {
    let out = run_prints(
        r#"
        interface Reader<out T> {
            fun read(): T
        }

        class NameReader : Reader<String> {
            override fun read(): String {
                return "ok"
            }
        }

        fun consume(reader: Reader<Any>): String {
            return reader.read().toString()
        }

        fun main() {
            val reader: Reader<String> = NameReader()
            println(consume(reader))
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_generic_in_variance_writer() {
    let out = run_prints(
        r#"
        interface Writer<in T> {
            fun write(value: T)
        }

        class Logger : Writer<Any> {
            var last: Any? = null

            override fun write(value: Any) {
                last = value
            }
        }

        fun emitInt(writer: Writer<Int>, value: Int): String {
            writer.write(value)
            return writer.toString()
        }

        fun main() {
            val logger = Logger()
            val writer: Writer<Int> = logger
            writer.write(7)
            println(logger.last)
        }
    "#,
    );
    assert_eq!(out, &["7"]);
}

#[test]
fn test_generic_nullable_type_parameter() {
    let out = run_prints(
        r#"
        fun <T> unwrapOrNone(value: T?): String {
            return value?.toString() ?: "none"
        }

        fun main() {
            println(unwrapOrNone("ok"))
            println(unwrapOrNone(null as String?))
            println(unwrapOrNone(0))
        }
    "#,
    );
    assert_eq!(out, &["ok", "none", "0"]);
}

#[test]
fn test_generic_class_stateful_update() {
    let out = run_prints(
        r#"
        class Store<T>(start: T) {
            private var value: T = start

            fun get(): T = value
            fun set(next: T) {
                value = next
            }
        }

        fun main() {
            val text = Store("a")
            val number = Store(10)
            text.set("b")
            number.set(11)
            println(text.get())
            println(number.get())
        }
    "#,
    );
    assert_eq!(out, &["b", "11"]);
}

#[test]
fn test_generic_function_with_multiple_return_types() {
    let out = run_prints(
        r#"
        fun <A, B> pairLabel(left: A, right: B): String {
            return left.toString() + ":" + right.toString()
        }

        fun main() {
            println(pairLabel(true, 1))
            println(pairLabel(2.2, "x"))
            println(pairLabel("k", false))
        }
    "#,
    );
    assert_eq!(out, &["true:1", "2.2:x", "k:false"]);
}

#[test]
fn test_generic_array_element_roundtrip() {
    let out = run_prints(
        r#"
        fun <T> firstAndLast(values: Array<T>): String {
            return values.first().toString() + "," + values.last().toString()
        }

        fun main() {
            println(firstAndLast(arrayOf(1, 2, 3)))
            println(firstAndLast(arrayOf("x", "y", "z")))
        }
    "#,
    );
    assert_eq!(out, &["1,3", "x,z"]);
}

#[test]
fn test_generic_typealias_and_alias_target() {
    let out = run_prints(
        r#"
        fun main() {
            val pair = Holder("left", "right")
            println(pair.parts())
        }
    "#,
    );
    assert_eq!(out, &["left:right"]);
}

#[test]
fn test_generic_local_type_inference_in_nested_scope() {
    let out = run_prints(
        r#"
        class Holder<T>(private val value: T) {
            fun value(): T = value
        }

        fun <T> describe(value: Holder<T>): String {
            return value.value().toString()
        }

        fun main() {
            val number = Holder(8)
            val text = Holder("gen")
            val inferred = describe(number)
            val direct = describe(text)
            println(inferred)
            println(direct)
        }
    "#,
    );
    assert_eq!(out, &["8", "gen"]);
}

#[test]
fn test_generic_star_projection_readonly() {
    let out = run_prints(
        r#"
        fun <T> consumeUnknown(values: Array<out T>): Int {
            if (values.size == 0) {
                return 0
            }
            return values.size
        }

        fun main() {
            val items: Array<Int> = arrayOf(1, 2, 3)
            println(consumeUnknown(items))
            println(consumeUnknown(arrayOf<String>()))
        }
    "#,
    );
    assert_eq!(out, &["3", "0"]);
}

#[test]
fn test_generic_function_returning_array_and_size() {
    let out = run_prints(
        r#"
        fun <T> toArray(left: T, right: T): Array<T> {
            return arrayOf(left, right)
        }

        fun main() {
            val numbers = toArray(2, 3)
            val words = toArray("a", "b")
            println(numbers.size)
            println(words.size)
            println(numbers[1] + words[1])
        }
    "#,
    );
    assert_eq!(out, &["2", "2", "3b"]);
}

#[test]
fn test_generic_function_with_three_comparable_values() {
    let out = run_prints(
        r#"
        fun <T : Comparable<T>> maxOfThree(a: T, b: T, c: T): T {
            return if (a > b && a > c) a else if (b > c) b else c
        }

        fun main() {
            println(maxOfThree(4, 9, 1))
            println(maxOfThree("alpha", "gamma", "beta"))
        }
    "#,
    );
    assert_eq!(out, &["9", "gamma"]);
}

#[test]
fn test_generic_collection_bridge_between_subtypes() {
    let out = run_prints(
        r#"
        fun <T> mergeInto(dest: MutableList<T>, first: T, second: T) {
            dest.add(first)
            dest.add(second)
        }

        fun main() {
            val data = mutableListOf<Any>()
            mergeInto(data, 1, "x")
            println(data.size)
            println(data[0])
            println(data[1])
        }
    "#,
    );
    assert_eq!(out, &["2", "1", "x"]);
}

#[test]
fn test_generic_function_accepts_nullable_bounded_any() {
    let out = run_prints(
        r#"
        fun <T : Any> ensureNotNull(value: T?): String {
            return value?.toString() ?: "missing"
        }

        fun main() {
            println(ensureNotNull(4))
            println(ensureNotNull("z"))
        }
    "#,
    );
    assert_eq!(out, &["4", "z"]);
}

#[test]
fn test_generic_function_reference_inference() {
    let out = run_prints(
        r#"
        fun <T> apply(value: T, op: (T) -> T): T {
            return op(value)
        }

        fun inc(value: Int): Int = value + 1
        fun shout(value: String): String = value.toUpperCase()

        fun main() {
            println(apply(2, ::inc))
            println(apply("ok", ::shout))
        }
    "#,
    );
    assert_eq!(out, &["3", "OK"]);
}

#[test]
fn test_generic_where_clause_dual_constraints() {
    let out = run_prints(
        r#"
        fun <T> compareAndMeasure(value: T): String
        where T : Comparable<T>, T : CharSequence {
            return value.length.toString() + ":" + if (value > value) "gt" else "eq"
        }

        fun main() {
            println(compareAndMeasure("abc"))
        }
    "#,
    );
    assert_eq!(out, &["3:eq"]);
}

#[test]
fn test_generic_typealias_builder_infers_type() {
    let out = run_prints(
        r#"
        typealias Factory<T> = () -> T

        fun <T> materialize(factory: Factory<T>): T {
            return factory()
        }

        fun main() {
            val number = materialize { 9 }
            val text = materialize { "go" }
            println(number)
            println(text)
        }
    "#,
    );
    assert_eq!(out, &["9", "go"]);
}

#[test]
fn test_generic_receiver_extension_with_secondary_projection() {
    let out = run_prints(
        r#"
        class Holder<T>(private val value: T) {
            fun value(): T = value
        }

        fun <T> Holder<T>.bind(other: T): String {
            return this.value().toString() + ":" + other.toString()
        }

        fun main() {
            val text = Holder("a")
            val numbers = Holder(4)
            println(text.bind("x"))
            println(numbers.bind(6))
        }
    "#,
    );
    assert_eq!(out, &["a:x", "4:6"]);
}

#[test]
fn test_generic_factory_function_with_variadic_tuple_emulation() {
    let out = run_prints(
        r#"
        fun <A, B> makePair(first: A, second: B): Array<Any?> {
            return arrayOf(first, second)
        }

        fun main() {
            val pair1 = makePair(1, "k")
            val pair2 = makePair(true, 3.2)
            println(pair1[0])
            println(pair1[1])
            println(pair2[0].toString() + pair2[1].toString())
        }
    "#,
    );
    assert_eq!(out, &["1", "k", "true3.2"]);
}

#[test]
fn test_generic_method_type_argument_overload_resolution() {
    let out = run_prints(
        r#"
        fun <T> format(value: T): String = value.toString()

        fun <T> format(value: T, prefix: String): String {
            return prefix + ":" + value
        }

        fun main() {
            println(format(9))
            println(format(9, "num"))
        }
    "#,
    );
    assert_eq!(out, &["9", "num:9"]);
}

#[test]
fn test_generic_covariant_readonly_collection_can_receive_concrete_subtype() {
    let out = run_prints(
        r#"
        fun <T> countValues(values: List<T>): Int {
            return values.size
        }

        fun main() {
            val ints: MutableList<Int> = mutableListOf(1, 2, 3)
            val numbers: List<Number> = ints
            println(countValues(numbers))
            println(countValues(ints))
        }
    "#,
    );
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_generic_recursive_data_structure_preserves_type() {
    let out = run_prints(
        r#"
        data class Node<T>(val value: T, val next: Node<T>? = null)

        fun <T> collect(values: Node<T>): String {
            var cursor: Node<T>? = values
            var out = ""
            while (cursor != null) {
                out += cursor.value.toString()
                cursor = cursor.next
                if (cursor != null) out += "-"
            }
            return out
        }

        fun main() {
            val chain = Node("a", Node("b", Node("c")))
            println(collect(chain))
        }
    "#,
    );
    assert_eq!(out, &["a-b-c"]);
}

#[test]
fn test_generic_class_with_shadowed_type_parameter() {
    let out = run_prints(
        r#"
        class Holder<T>(val value: T) {
            fun <R> map(transform: (T) -> R): Holder<R> {
                return Holder(transform(value))
            }
        }

        fun main() {
            val holder = Holder("7")
            val number = holder.map { it.toInt() }
            val text = holder.map { it + it }
            println(number.value + 1)
            println(text.value)
        }
    "#,
    );
    assert_eq!(out, &["8", "77"]);
}

#[test]
fn test_generic_two_way_variance_contract() {
    let out = run_prints(
        r#"
        interface Converter<in S, out T> {
            fun convert(value: S): T
        }

        class StringToInt : Converter<String, Number> {
            override fun convert(value: String): Number = value.length
        }

        fun emit(any: Converter<CharSequence, Number>, value: CharSequence): String {
            return any.convert(value).toString()
        }

        fun main() {
            val converter: Converter<Any, Number> = StringToInt()
            println(emit(converter, "abc"))
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_generic_projection_of_list_producer_consumer() {
    let out = run_prints(
        r#"
        interface Producer<out T> {
            fun produce(): T
        }

        interface Consumer<in T> {
            fun consume(value: T)
        }

        class StringProducer : Producer<String> {
            var value = "go"
            override fun produce(): String = value
        }

        class AnyConsumer : Consumer<Any> {
            var last: Any? = null
            override fun consume(value: Any) { last = value }
            fun seen(): String = last.toString()
        }

        fun pipe(source: Producer<String>, sink: Consumer<CharSequence>) {
            sink.consume(source.produce())
        }

        fun main() {
            val producer: Producer<String> = StringProducer()
            val consumer = AnyConsumer()
            pipe(producer, consumer)
            println(consumer.seen())
        }
    "#,
    );
    assert_eq!(out, &["go"]);
}

#[test]
fn test_generic_where_clause_with_number_and_serializable() {
    let out = run_prints(
        r#"
        import java.io.Serializable

        fun <T> describe(value: T): String
        where T : Number, T : Serializable {
            return value.toString()
        }

        fun main() {
            println(describe(12))
            println(describe(3.4))
        }
    "#,
    );
    assert_eq!(out, &["12", "3.4"]);
}

#[test]
fn test_generic_function_rejects_incompatible_constraints_by_type_inference() {
    let out = run_prints(
        r#"
        fun <T> pairSize(left: T, right: T): Int {
            return 2
        }

        class Item

        fun main() {
            println(pairSize(Item(), Item()))
            val left = Item()
            val right = Item()
            println(pairSize(left, right))
        }
    "#,
    );
    assert_eq!(out, &["2", "2"]);
}

#[test]
fn test_generic_nested_generic_function_chain() {
    let out = run_prints(
        r#"
        fun <T> wrap(value: T): Box<T> = Box(value)
        class Box<T>(val value: T)

        fun <T, R> wrapChain(value: T, op: (T) -> R): Box<R> {
            return wrap(op(value))
        }

        fun main() {
            val boxed = wrapChain(3, { it + 1 })
            val text = wrapChain("ok", { it + "!" })
            println(boxed.value)
            println(text.value)
        }
    "#,
    );
    assert_eq!(out, &["4", "ok!"]);
}

#[test]
fn test_generic_map_projection_of_collections() {
    let out = run_prints(
        r#"
        fun <T> toStringMap(values: Map<String, T>): Map<String, String> {
            return values.mapValues { it.value.toString() }
        }

        fun main() {
            val input: Map<String, Int> = mapOf("a" to 1, "b" to 2)
            val projected = toStringMap(input)
            println(projected["a"])
            println(projected["b"])
        }
    "#,
    );
    assert_eq!(out, &["1", "2"]);
}
