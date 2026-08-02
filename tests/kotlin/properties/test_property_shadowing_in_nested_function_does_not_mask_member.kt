// vybe-test: kotlin/properties/test_property_shadowing_in_nested_function_does_not_mask_member
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder(var value: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder("member")
            fun readLabel(value: String): String {
                return holder.value + "-" + value
            }
            __check((readLabel("arg")).toString(), "member-arg")
            holder.value = "next"
            __check((readLabel("arg")).toString(), "next-arg")
        }
