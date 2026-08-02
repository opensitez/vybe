// vybe-test: kotlin/variance/test_variance_generics_with_projection_in_class
// origin: languages/kotlin/tests/kotlin/test_variance.rs

class Box<T>(val value: T)
        fun printBox(values: Box<out Any>) {
            println(values.value)
        }
        fun main() {
            printBox(Box("hello"))
            printBox(Box(99))
        }

