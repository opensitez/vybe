kotlin_run_test!(
    test_top_level_reference_call,
    r#"
        fun square(v: Int) = v * v
        fun main() {
            val f = ::square
            println(f(7))
        }
    "#,
    &["49"]
);

kotlin_run_test!(
    test_top_level_reference_with_multiple_args,
    r#"
        fun add(a: Int, b: Int) = a + b
        fun main() {
            val f = ::add
            println(f(3, 4))
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_top_level_reference_as_higher_order_value,
    r#"
        fun identity(v: Int) = v + 10
        fun apply(value: Int, fn: (Int) -> Int): Int = fn(value)
        fun main() {
            println(apply(3, ::identity))
        }
    "#,
    &["13"]
);

kotlin_run_test!(
    test_constructor_reference,
    r#"
        class Box(val value: Int)
        fun main() {
            val ctor = ::Box
            val x = ctor(4)
            println(x.value)
        }
    "#,
    &["4"]
);

kotlin_run_test!(
    test_constructor_reference_in_collection,
    r#"
        class Tag(val label: String)
        fun main() {
            val names = listOf("a", "bb").map(::Tag).map { it.label }
            println(names.joinToString(","))
        }
    "#,
    &["a,bb"]
);

kotlin_run_test!(
    test_member_function_reference,
    r#"
        class Counter(val step: Int) {
            fun plus(v: Int) = v + step
        }
        fun main() {
            val add = Counter(5)::plus
            println(add(7))
        }
    "#,
    &["12"]
);

kotlin_run_test!(
    test_bound_member_function_reference,
    r#"
        class Greeter(val name: String) {
            fun hello(prefix: String) = "$prefix$name"
        }
        fun main() {
            val g = Greeter("k")
            val hi = g::hello
            println(hi("x"))
        }
    "#,
    &["xk"]
);

kotlin_run_test!(
    test_member_property_reference,
    r#"
        class User(val id: Int)
        fun main() {
            val readId = User::id
            println(readId(User(9)))
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_bound_member_property_reference,
    r#"
        class User(val id: Int)
        fun main() {
            val u = User(11)
            val read = u::id
            println(read())
        }
    "#,
    &["11"]
);

kotlin_run_test!(
    test_standard_extension_property_reference,
    r#"
        fun main() {
            val ref = String::length
            println(ref("kotlin"))
        }
    "#,
    &["6"]
);

kotlin_run_test!(
    test_standard_extension_function_reference,
    r#"
        fun main() {
            val trimRef: (String) -> String = String::trim
            println(trimRef("  a "))
        }
    "#,
    &["a"]
);

kotlin_run_test!(
    test_instance_method_reference_in_map,
    r#"
        class Item(val value: String)
        fun main() {
            val items = listOf(Item("a"), Item("b"))
            val labels = items.map(Item::value).joinToString("|")
            println(labels)
        }
    "#,
    &["a|b"]
);

kotlin_run_test!(
    test_bound_method_reference_with_local_value,
    r#"
        class Holder(val tag: String) {
            fun emit() = tag
        }
        fun main() {
            val h = Holder("ok")
            val f = h::emit
            println(f())
        }
    "#,
    &["ok"]
);

kotlin_run_test!(
    test_unbound_property_reference_in_sorting,
    r#"
        data class Node(val score: Int)
        fun main() {
            val nodes = listOf(Node(2), Node(1), Node(3))
            val out = nodes.sortedBy(Node::score).joinToString(",") { it.score.toString() }
            println(out)
        }
    "#,
    &["1,2,3"]
);

kotlin_run_test!(
    test_extension_to_higher_order,
    r#"
        fun String.prefixWith(prefix: String) = prefix + this
        fun transform(values: List<String>, fn: (String) -> String): String =
            values.joinToString(",") { fn(it) }

        fun main() {
            println(transform(listOf("a", "b"), String::prefixWith("x"))
        }
    "#,
    &["xa,xb"]
);

kotlin_run_test!(
    test_reference_to_function_with_nullable_receiver,
    r#"
        fun main() {
            val pick = (String?)::orEmpty
            println(pick(null))
            println(pick("x"))
        }
    "#,
    &["", "x"]
);

kotlin_run_test!(
    test_reference_to_java_static_like,
    r#"
        fun main() {
            val fromInt = Int::toString
            println(fromInt(5))
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_reference_through_lambda_map,
    r#"
        fun shout(x: Int) = x + 1
        fun main() {
            val out = listOf(1, 2, 3).map(::shout).joinToString(";")
            println(out)
        }
    "#,
    &["2;3;4"]
);

kotlin_run_test!(
    test_reference_to_instance_of_nested_object,
    r#"
        class Box {
            inner class Inner {
                fun value(v: Int) = v + 1
            }
        }
        fun main() {
            val f = Box().Inner()::value
            println(f(8))
        }
    "#,
    &["9"]
);

kotlin_run_test!(
    test_reference_to_list_size_property,
    r#"
        fun main() {
            val readSize = List<Int>::size
            println(readSize(listOf(1, 2, 3)))
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_reference_to_map_getter,
    r#"
        fun main() {
            val read = Map<String, Int>::get
            val map = mapOf("x" to 7)
            println(read(map, "x"))
            println(read(map, "y"))
        }
    "#,
    &["7", "null"]
);

kotlin_run_test!(
    test_reference_to_boolean_extension_function,
    r#"
        val check: (Boolean) -> String = Boolean::toString

        fun main() {
            println(check(true))
            println(check(false))
        }
    "#,
    &["true", "false"]
);

kotlin_run_test!(
    test_reference_to_math_abs_function,
    r#"
        import kotlin.math.abs
        fun main() {
            val f = ::abs
            println(f(-10))
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_bound_property_reference_from_val_instance,
    r#"
        class Counter {
            val label: String = "c"
        }

        fun main() {
            val counter = Counter()
            val labelRef = counter::label
            println(labelRef())
        }
    "#,
    &["c"]
);

kotlin_run_test!(
    test_reference_to_mutable_property_on_instance,
    r#"
        class Holder {
            var value: Int = 0
        }

        fun main() {
            val holder = Holder()
            holder.value = 8
            val read = holder::value
            holder.value = 1
            val read2 = Holder::value
            println(read())
            println(read2(holder))
        }
    "#,
    &["1", "1"]
);

kotlin_run_test!(
    test_reference_to_constructor_with_args,
    r#"
        class Pair(val left: Int, val right: String)

        fun main() {
            val make = ::Pair
            val item = make(2, "x")
            println(item.left)
            println(item.right)
        }
    "#,
    &["2", "x"]
);

kotlin_run_test!(
    test_reference_to_member_function_with_receiver,
    r#"
        class Word(val value: String) {
            fun upper(): String = value.uppercase()
        }

        fun main() {
            val op = Word::upper
            println(op(Word("abc")))
        }
    "#,
    &["ABC"]
);

kotlin_run_test!(
    test_reference_to_nested_member_function,
    r#"
        class Board {
            class Cell {
                fun mark(v: String): String = "[$v]"
            }
        }

        fun main() {
            val label = Board.Cell::mark
            println(label(Board.Cell(), "x"))
        }
    "#,
    &["[x]"]
);

kotlin_run_test!(
    test_reference_in_map_chain_with_function_reference,
    r#"
        class Item(val value: Int)

        fun main() {
            val items = listOf(Item(3), Item(7), Item(9))
            val refs = items.map(Item::value).map { it * 2 }
            println(refs.joinToString("|"))
        }
    "#,
    &["6|14|18"]
);

kotlin_run_test!(
    test_reference_to_extension_receiver_function,
    r#"
        fun String.surround(left: String, right: String): String = left + this + right

        fun main() {
            val ref: (String, String, String) -> String = String::surround
            println(ref("k", "<", ">"))
        }
    "#,
    &["<k>"]
);
