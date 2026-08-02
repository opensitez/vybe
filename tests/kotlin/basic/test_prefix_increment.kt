// vybe-test: kotlin/basic/test_prefix_increment
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 1
            __check((++count).toString(), "2")
            __check((--count).toString(), "1")
        }
