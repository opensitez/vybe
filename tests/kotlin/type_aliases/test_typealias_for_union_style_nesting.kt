// vybe-test: kotlin/type_aliases/test_typealias_for_union_style_nesting
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias Name = String
        typealias Named = Name

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Named = "x"
            __check((value).toString(), "x")
        }
