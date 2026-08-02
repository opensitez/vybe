// vybe-test: kotlin/member_references/test_reference_to_map_getter
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val read = Map<String, Int>::get
            val map = mapOf("x" to 7)
            __check((read(map, "x")).toString(), "7")
            __check((read(map, "y")).toString(), "null")
        }
