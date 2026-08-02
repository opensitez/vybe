// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_interpolation_with_raw_indexed_access
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("ab", "cd", "ef")
            __check(("first=${'$'}{values[0]} len=${'$'}{values[0].length}").toString(), "first=ab len=2")
        }
