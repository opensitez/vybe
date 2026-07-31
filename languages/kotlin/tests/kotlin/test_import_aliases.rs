kotlin_run_test!(
    test_import_alias_function_alias_usage,
    r#"
        import kotlin.math.absoluteValue as absValue
        fun main() {
            println(absValue(-11))
        }
    "#,
    &["11"]
);

kotlin_run_test!(
    test_import_alias_type_aliasing,
    r#"
        import kotlin.collections.HashMap as MapAlias
        fun main() {
            val map = MapAlias<String, Int>()
            map["a"] = 1
            println(map["a"])
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_import_alias_conflict_with_local_type,
    r#"
        import kotlin.collections.List as KotlinList
        fun main() {
            val local: KotlinList<Int> = listOf(1, 2, 3)
            println(local.size)
        }
    "#,
    &["3"]
);

kotlin_run_test!(
    test_import_alias_chain_of_calls,
    r#"
        import kotlin.math.max as takeMax
        import kotlin.math.min as takeMin

        fun main() {
            println(takeMax(3, 9))
            println(takeMin(3, 9))
        }
    "#,
    &["9", "3"]
);

kotlin_run_test!(
    test_import_alias_abs_reference,
    r#"
        import kotlin.math.absoluteValue as absAlias

        fun main() {
            println(absAlias(-5))
            println(absAlias(6))
        }
    "#,
    &["5", "6"]
);

kotlin_run_test!(
    test_import_alias_star_kept_local_shadowing,
    r#"
        import kotlin.collections.*

        fun main() {
            val mutableList = mutableListOf("a", "b")
            println(mutableList.joinToString("/"))
            println(listOf(1, 2).size)
        }
    "#,
    &["a/b", "2"]
);

kotlin_run_test!(
    test_import_alias_local_override_of_import,
    r#"
        import kotlin.math.sqrt as squareRoot

        fun main() {
            fun squareRoot(x: Int): Int = x * x
            val f: (Double) -> Double = kotlin.math::sqrt
            println(squareRoot(3))
            println(f(4.0).toInt())
        }
    "#,
    &["9", "2"]
);

kotlin_run_test!(
    test_import_alias_multiple_namespaces,
    r#"
        import java.lang.StringBuilder as KotlinBuilder
        import java.util.StringTokenizer as Tokenizer

        fun main() {
            val builder = KotlinBuilder()
            builder.append("a").append("b")
            println(builder.toString())
            val tokenizer = Tokenizer("x,y", ",")
            var tokens = 0
            while (tokenizer.hasMoreTokens()) {
                tokenizer.nextToken()
                tokens += 1
            }
            println(tokens)
        }
    "#,
    &["ab", "2"]
);

kotlin_run_test!(
    test_import_alias_with_class_and_function_usage,
    r#"
        import kotlin.math.round as roundFunction

        fun normalize(v: Double): Int {
            return roundFunction(v).toInt()
        }

        fun main() {
            println(normalize(2.4))
            println(roundFunction(2.9).toInt())
        }
    "#,
    &["2", "3"]
);

kotlin_run_test!(
    test_import_alias_for_nested_path,
    r#"
        import kotlin.collections.ArrayList as KotlinIntList

        fun main() {
            val values: KotlinIntList<Int> = KotlinIntList()
            values.add(5)
            values.add(6)
            println(values[0] + values[1])
        }
    "#,
    &["11"]
);

kotlin_run_test!(
    test_import_alias_for_function_value,
    r#"
        import kotlin.math.max as pickMax

        fun main() {
            println(pickMax(3, 10))
            println(pickMax(-1, 2))
        }
    "#,
    &["10", "2"]
);

kotlin_run_test!(
    test_import_alias_isolated_scopes,
    r#"
        import kotlin.math.abs as absValue

        fun score(v: Int): Int {
            val absValue = { x: Int -> x * x }
            return absValue(v)
        }

        fun main() {
            println(score(-3))
            println(absValue(-3))
        }
    "#,
    &["9", "3"]
);
