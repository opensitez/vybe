// vybe-test: kotlin/type_inference/test_type_inference_with_class_hierarchy
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

open class Base
        class Child : Base()
        fun id(base: Base): Base = base
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = id(Child())
            __check((value is Child).toString(), "true")
        }
