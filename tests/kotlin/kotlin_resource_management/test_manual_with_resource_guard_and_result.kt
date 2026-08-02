// vybe-test: kotlin/kotlin_resource_management/test_manual_with_resource_guard_and_result
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Counter : AutoCloseable {
            var total = 0
            override fun close() {
                total += 10
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            c.use {
                it.total = 5
                __check((it.total).toString(), "5")
            }
            __check((c.total).toString(), "15")
        }
