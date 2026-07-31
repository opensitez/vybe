kotlin_run_test!(
    test_upper_bound_comparable_generic_function,
    r#"
        fun <T : Comparable<T>> maxValue(a: T, b: T): T = if (a >= b) a else b

        fun main() {
            println(maxValue(3, 7))
            println(maxValue("a", "b"))
        }
    "#,
    &["7", "b"]
);

kotlin_run_test!(
    test_multiple_bounds_for_numeric_chain,
    r#"
        fun <T> sumPositive(a: T, b: T): Double where T : Number, T : Comparable<T> {
            return a.toDouble() + b.toDouble()
        }

        fun main() {
            println(sumPositive(2, 3))
            println(sumPositive(1.5, 2.5))
        }
    "#,
    &["5", "4"]
);

kotlin_run_test!(
    test_generic_extension_with_reified_like_bound_checks,
    r#"
        class Holder<T : Any>(val value: T)

        fun <T : Any> describe(value: T): String {
            return value::class.simpleName ?: ""
        }

        fun main() {
            println(describe(Holder(1)))
            println(describe("x").length)
        }
    "#,
    &["Holder", "1"]
);

kotlin_run_test!(
    test_invariant_collection_bound,
    r#"
        fun <T> firstOf(list: List<T>): T {
            return list[0]
        }

        fun main() {
            println(firstOf(listOf(1, 2, 3)))
            println(firstOf(listOf("a", "b")))
        }
    "#,
    &["1", "a"]
);

kotlin_run_test!(
    test_covariant_return_generic_class,
    r#"
        open class Base
        class Child : Base()

        class Box<T : Base>(val payload: T)

        fun <T : Base> identity(box: Box<T>): T = box.payload

        fun main() {
            val c = Box(Child())
            val base: Base = identity(c)
            println(base is Child)
        }
    "#,
    &["true"]
);

kotlin_run_test!(
    test_where_clause_in_generic_function,
    r#"
        fun <T> toText(value: T): String where T : Any {
            return value.toString()
        }

        fun main() {
            println(toText("ok"))
            println(toText(5))
        }
    "#,
    &["ok", "5"]
);

kotlin_run_test!(
    test_generic_with_callable_reference_bound,
    r#"
        fun <T : Number> describe(value: T): String {
            return value.toString() + value.toInt().toString()
        }

        fun main() {
            val fn: (Int) -> String = ::describe
            println(fn(3))
        }
    "#,
    &["33"]
);

kotlin_run_test!(
    test_type_parameter_with_enum_constraint,
    r#"
        interface Named { fun name(): String }

        enum class Source : Named { A { override fun name() = "a" }, B { override fun name() = "b" } }

        fun <T> label(item: T): String where T : Enum<T>, T : Named {
            return item.name()
        }

        fun main() {
            println(label(Source.A))
            println(label(Source.B))
        }
    "#,
    &["a", "b"]
);

kotlin_run_test!(
    test_generic_infer_from_assignment,
    r#"
        class Box<T>(val value: T)

        fun <T> firstOrDefault(value: T?): T {
            return value ?: throw IllegalStateException("missing")
        }

        fun main() {
            val v: String? = "k"
            val b = Box(firstOrDefault(v))
            println(b.value)
        }
    "#,
    &["k"]
);

kotlin_run_test!(
    test_generic_constraint_chain,
    r#"
        class NumericBox<T>(val value: T) where T : Number, T : Comparable<T>

        fun <T> maxBox(a: NumericBox<T>, b: NumericBox<T>): T where T : Number, T : Comparable<T> {
            return if (a.value >= b.value) a.value else b.value
        }

        fun main() {
            val a = NumericBox(3)
            val b = NumericBox(7)
            println(maxBox(a, b))
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_generic_typealias_with_upper_bound,
    r#"
        typealias ComparableNumber<T> = T where T : Comparable<T>, T : Number

        fun pick(a: ComparableNumber<Int>, b: ComparableNumber<Int>): Int = if (a > b) a else b

        fun main() {
            println(pick(10, 4))
        }
    "#,
    &["10"]
);
