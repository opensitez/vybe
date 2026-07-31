kotlin_run_test!(
    test_reified_type_check_for_int,
    r#"
        inline fun <reified T> isType(value: Any?): String {
            return if (value is T) "yes" else "no"
        }

        fun main() {
            println(isType<Int>(3))
            println(isType<String>(3))
        }
    "#,
    &["yes", "no"]
);

kotlin_run_test!(
    test_reified_nullable_check,
    r#"
        inline fun <reified T> safeCast(value: Any?): Boolean = value is T

        fun main() {
            val a: String? = null
            val b: String? = "x"
            println(safeCast<String?>(a))
            println(safeCast<String?>(b))
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_reified_generic_name,
    r#"
        inline fun <reified T> typeName(): String = T::class.simpleName ?: "unknown"

        fun main() {
            println(typeName<Int>())
            println(typeName<List<String>>())
        }
    "#,
    &["Int", "List"]
);

kotlin_run_test!(
    test_reified_as_cast_result,
    r#"
        inline fun <reified T> asOrNull(value: Any?): String {
            val cast = value as? T
            return if (cast == null) "none" else "has"
        }

        fun main() {
            println(asOrNull<String>("kotlin"))
            println(asOrNull<String>(8))
        }
    "#,
    &["has", "none"]
);

kotlin_run_test!(
    test_reified_list_type_check,
    r#"
        inline fun <reified T> hasStrings(values: List<Any>): String {
            return if (values.all { it is T }) "all" else "some"
        }

        fun main() {
            println(hasStrings<String>(listOf("a", "b", "c")))
            println(hasStrings<Int>(listOf("a", 1, "c")))
        }
    "#,
    &["all", "some"]
);

kotlin_run_test!(
    test_reified_generic_identity,
    r#"
        inline fun <reified T> sameType(a: T, b: T): String = if (a::class == b::class) "same" else "diff"

        fun main() {
            println(sameType(1, 2))
            println(sameType("a", "b"))
        }
    "#,
    &["same", "same"]
);

kotlin_run_test!(
    test_reified_list_of_types,
    r#"
        inline fun <reified T> describe(values: List<T>): String = values::class.simpleName ?: ""

        fun main() {
            println(describe(listOf(1, 2, 3)))
            println(describe(listOf("a", "b")))
        }
    "#,
    &["ArrayList", "ArrayList"]
);

kotlin_run_test!(
    test_reified_pair_check,
    r#"
        inline fun <reified T> isPair(value: Any?): String = if (value is Pair<T, T>) "pair" else "not-pair"

        fun main() {
            println(isPair<Int>(Pair(1, 2)))
            println(isPair<String>(Pair(1, "x")))
        }
    "#,
    &["pair", "not-pair"]
);
