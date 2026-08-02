// vybe-test: kotlin/throwing_recovery/test_catch_ordering_specificity
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: RuntimeException) {
                println("runtime")
            } catch (e: Exception) {
                println("general")
            }
        }

