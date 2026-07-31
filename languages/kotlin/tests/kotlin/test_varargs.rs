kotlin_run_test!(
    test_vararg_int_sum,
    r#"
        fun sumAll(vararg values: Int): Int = values.sum()

        fun main() {
            println(sumAll(1, 2, 3, 4))
        }
    "#,
    &["10"]
);

kotlin_run_test!(
    test_vararg_string_join,
    r#"
        fun join(base: String, vararg values: String): String =
            base + values.joinToString(":")

        fun main() {
            println(join("x", "a", "b", "c"))
        }
    "#,
    &["xa:b:c"]
);

kotlin_run_test!(
    test_vararg_with_named_arguments,
    r##"
        fun build(prefix: String, suffix: String = "#", vararg values: String): String {
            return prefix + values.joinToString(suffix)
        }

        fun main() {
            println(build("v", ";", "one", "two"))
        }
    "##,
    &["vone;two"]
);

kotlin_run_test!(
    test_vararg_spread_from_array,
    r#"
        fun maxOfAll(vararg values: Int): Int = values.maxOrNull() ?: 0

        fun main() {
            val base = intArrayOf(4, 1, 8)
            println(maxOfAll(*base))
        }
    "#,
    &["8"]
);

kotlin_run_test!(
    test_vararg_empty_behavior,
    r#"
        fun count(vararg values: Int): Int = values.size

        fun main() {
            println(count())
        }
    "#,
    &["0"]
);

kotlin_run_test!(
    test_vararg_in_class_method,
    r#"
        class Collector {
            fun collect(vararg items: String): String = items.joinToString(",")
        }

        fun main() {
            val c = Collector()
            println(c.collect("x", "y"))
        }
    "#,
    &["x,y"]
);

kotlin_run_test!(
    test_vararg_any_type,
    r#"
        fun describe(vararg values: Any): String = values.joinToString("|") { it.toString() }

        fun main() {
            println(describe("a", 2, true))
        }
    "#,
    &["a|2|true"]
);

kotlin_run_test!(
    test_vararg_extension_function,
    r#"
        fun String.wrapAll(vararg values: String): String = values.joinToString(this, prefix = "<", postfix = ">")

        fun main() {
            println(",".wrapAll("x", "y", "z"))
        }
    "#,
    &["<x,y,z>"]
);

kotlin_run_test!(
    test_vararg_array_list_conversion,
    r#"
        fun asList(prefix: String, vararg values: Int): String {
            val list = values.toList().map { it.toString() }.joinToString(prefix = prefix, separator = "-")
            return list
        }

        fun main() {
            println(asList("a", 1, 2, 3))
        }
    "#,
    &["1-2-3"]
);

kotlin_run_test!(
    test_vararg_nullable_elements,
    r#"
        fun read(values: vararg value: String?): String {
            return values.joinToString(";") { it ?: "nil" }
        }

        fun main() {
            println(read("x", null, "z"))
        }
    "#,
    &["x;nil;z"]
);

kotlin_run_test!(
    test_vararg_used_alongside_single_list_argument,
    r#"
        fun append(base: String, separator: String, vararg values: String): String {
            return base + values.joinToString(separator)
        }

        fun main() {
            println(append("v", ".", "a", "b", "c"))
        }
    "#,
    &["v.a.b.c"]
);
