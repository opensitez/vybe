// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_uses_value_once_per_instance
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var count = 0
            class Holder {
                val value by lazy { count++
"x" }
            }
            val first = Holder()
            val second = Holder()
            __check((first.value).toString(), "x")
            __check((first.value).toString(), "x")
            __check((second.value).toString(), "x")
            __check((count).toString(), "2")
        }
