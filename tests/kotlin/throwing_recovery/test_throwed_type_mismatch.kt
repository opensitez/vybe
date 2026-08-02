// vybe-test: kotlin/throwing_recovery/test_throwed_type_mismatch
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun main() {
            try {
                val value: Any = "text"
                val num = value as Int
                println(num)
            } catch (e: ClassCastException) {
                println("cast")
            }
        }

