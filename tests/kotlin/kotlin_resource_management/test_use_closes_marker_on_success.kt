// vybe-test: kotlin/kotlin_resource_management/test_use_closes_marker_on_success
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Marker : AutoCloseable {
            var closed = false
            override fun close() {
                closed = true
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val marker = Marker()
            marker.use {
                __check(("open").toString(), "open")
            }
            __check((marker.closed).toString(), "true")
        }
