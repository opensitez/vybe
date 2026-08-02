// vybe-test: kotlin/kotlin_resource_management/test_try_finally_always_runs_close
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Marker {
            var closed = false
            fun close() { closed = true }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val marker = Marker()
            try {
                __check(("inside").toString(), "inside")
            } finally {
                marker.close()
            }
            __check((marker.closed).toString(), "true")
        }
