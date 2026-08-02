// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_default_initializes_once
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var initCount = 0
            val x: Int by lazy {
                initCount++
                5 + 3
            }
            __check((initCount).toString(), "0")
            __check((x).toString(), "8")
            __check((x).toString(), "8")
            __check((initCount).toString(), "1")
        }
