// vybe-test: kotlin/member_references/test_extension_to_higher_order
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun String.prefixWith(prefix: String) = prefix + this
        fun transform(values: List<String>, fn: (String) -> String): String =
            values.joinToString(",") { fn(it) }

        fun main() {
            println(transform(listOf("a", "b"), String::prefixWith("x"))
        }

