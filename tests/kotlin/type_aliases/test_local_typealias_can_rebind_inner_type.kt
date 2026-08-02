// vybe-test: kotlin/type_aliases/test_local_typealias_can_rebind_inner_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            typealias LocalMap = Map<String, Int>
            val values: LocalMap = mapOf("x" to 10)
            __check((values["x"]).toString(), "10")
        }
