// vybe-test: kotlin/smart_casts/test_type_test_with_boolean_and_and_guard
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

open class Base
        class Child : Base()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Child()
            __check((value is Base && value is Child).toString(), "true")
            __check((!(value is Child && value is String)).toString(), "true")
        }
