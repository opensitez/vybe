// vybe-test: kotlin/throwing_recovery/test_throwing_multiple_fails
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun fail(kind: Int) {
            when (kind) {
                1 -> throw IllegalArgumentException("a")
                2 -> throw IllegalStateException("b")
                else -> throw Exception("c")
            }
        }
        fun main() {
            for (i in 1..3) {
                try {
                    fail(i)
                } catch (e: Exception) {
                    println(e::class.java.simpleName)
                }
            }
        }

