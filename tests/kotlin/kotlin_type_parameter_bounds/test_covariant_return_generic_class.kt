// vybe-test: kotlin/kotlin_type_parameter_bounds/test_covariant_return_generic_class
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

open class Base
        class Child : Base()

        class Box<T : Base>(val payload: T)

        fun <T : Base> identity(box: Box<T>): T = box.payload

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Box(Child())
            val base: Base = identity(c)
            __check((base is Child).toString(), "true")
        }
