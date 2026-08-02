// vybe-test: kotlin/type_casts/test_sibling_downcast_failure_path
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

open class Shape
        class Circle : Shape()
        class Square : Shape()

        fun main() {
            val value: Shape = Square()
            try {
                val casted = value as Circle
                println(casted == null)
            } catch (e: Exception) {
                println("caught")
            }
        }

