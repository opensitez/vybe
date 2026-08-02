// vybe-test: kotlin/scope_shadowing/test_property_shadowing_in_nested_class
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

open class Base(val value: String)
        class Holder(overrideValue: String) : Base("base") {
            val value = overrideValue
            fun show(): String {
                return value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder("inner")
            __check((holder.show()).toString(), "inner")
            __check(((holder as Base).value).toString(), "base")
        }
