kotlin_run_test!(
    test_import_alias_for_function,
    r#"
        import kotlin.math.abs as kotlinAbs

        fun main() {
            println(kotlinAbs(-7))
        }
    "#,
    &["7"]
);

kotlin_run_test!(
    test_import_alias_for_class,
    r#"
        import kotlin.collections.HashMap as KMap

        fun main() {
            val map: KMap<String, Int> = KMap()
            map["a"] = 1
            println(map["a"])
        }
    "#,
    &["1"]
);

kotlin_run_test!(
    test_importing_multiple_aliases,
    r#"
        import kotlin.math.max as takeMax
        import kotlin.math.min as takeMin

        fun main() {
            println(takeMax(3, 7))
            println(takeMin(3, 7))
        }
    "#,
    &["7", "3"]
);

kotlin_run_test!(
    test_alias_function_with_same_name_shadowing,
    r#"
        import kotlin.collections.joinToString as join

        fun main() {
            val text = listOf("a", "b", "c").let { join(it, "/") }
            println(text)
        }
    "#,
    &["a/b/c"]
);

kotlin_run_test!(
    test_package_import_star_behavior,
    r#"
        import kotlin.math.*

        fun main() {
            println(sin(0.0) == 0.0)
            println(cos(0.0) == 1.0)
        }
    "#,
    &["true", "true"]
);

kotlin_run_test!(
    test_alias_type_parameterized_function,
    r#"
        import kotlin.collections.sortedBy as bySort

        fun main() {
            val values = listOf("z", "aa", "bbb")
            val sorted = values.bySort { it.length }
            println(sorted.joinToString(","))
        }
    "#,
    &["z,aa,bbb"]
);

kotlin_run_test!(
    test_import_alias_keeps_call_site_clear,
    r#"
        import kotlin.math.max as m

        fun main() {
            println(m(10, 2))
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_multiple_aliases_for_same_symbol_family,
    r#"
        import kotlin.collections.joinToString as joinA
        import kotlin.collections.joinToString as joinB

        fun main() {
            val text = listOf(1, 2).joinA(",")
            println(text)
            println(joinB(listOf(3, 4), "+"))
        }
    "#,
    &["1,2", "3+4"]
);

kotlin_run_test!(
    test_import_alias_for_nested_type_reference,
    r#"
        import kotlin.collections.List as KList

        fun main() {
            val values: KList<Int> = listOf(4, 5, 6)
            println(values.sum())
        }
    "#,
    &["15"]
);

kotlin_run_test!(
    test_aliased_local_class_name_collision,
    r#"
        import kotlin.collections.HashSet as Bucket

        fun main() {
            val left = Bucket<Int>()
            left.add(1)
            left.add(2)
            println(left.size)
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_import_alias_in_nested_expression,
    r#"
        import kotlin.collections.map as asMap

        fun main() {
            val out = listOf(1, 2, 3).asMap { it * it }
            println(out.joinToString(","))
        }
    "#,
    &["1,4,9"]
);
