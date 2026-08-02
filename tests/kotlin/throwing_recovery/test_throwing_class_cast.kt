// vybe-test: kotlin/throwing_recovery/test_throwing_class_cast
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun main() {
            val x: Any = 10
            try {
                val y = x as String
                println(y)
            } catch (e: ClassCastException) {
                println("class-cast")
            }
        }

