// vybe-test: kotlin/throwing_recovery/test_throwing_multiple_errors
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun explode(i: Int) {
            when (i) {
                0 -> throw IllegalArgumentException("zero")
                1 -> throw IllegalStateException("state")
                else -> println("ok")
            }
        }
        fun main() {
            try {
                explode(0)
            } catch (e: Exception) {
                println(e::class.java.simpleName)
            }
        }

