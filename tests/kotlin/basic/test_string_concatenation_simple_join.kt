// vybe-test: kotlin/basic/test_string_concatenation_simple_join
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
val s = "Hello " + "World"
            __check((s).toString(), "Hello World")
        }
