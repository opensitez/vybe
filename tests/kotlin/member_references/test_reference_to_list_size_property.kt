// vybe-test: kotlin/member_references/test_reference_to_list_size_property
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val readSize = List<Int>::size
            __check((readSize(listOf(1, 2, 3))).toString(), "3")
        }
