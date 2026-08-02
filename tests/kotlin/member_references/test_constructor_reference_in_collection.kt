// vybe-test: kotlin/member_references/test_constructor_reference_in_collection
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Tag(val label: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val names = listOf("a", "bb").map(::Tag).map { it.label }
            __check((names.joinToString(",")).toString(), "a,bb")
        }
