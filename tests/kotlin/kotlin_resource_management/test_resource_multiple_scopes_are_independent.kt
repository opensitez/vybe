// vybe-test: kotlin/kotlin_resource_management/test_resource_multiple_scopes_are_independent
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Flag : AutoCloseable {
            var closeCount = 0
            override fun close() { closeCount += 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Flag()
            val second = Flag()
            first.use { }
            second.use { }
            __check((first.closeCount).toString(), "1")
            __check((second.closeCount).toString(), "1")
        }
