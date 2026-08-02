// vybe-test: kotlin/kotlin_resource_management/test_resource_acquired_before_body_and_closed_after
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Holder : AutoCloseable {
            var active = false
            override fun close() { active = false }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var h: Holder? = null
            Holder().use {
                it.active = true
                h = it
                __check((it.active).toString(), "true")
            }
            __check((h?.active ?: false).toString(), "false")
        }
