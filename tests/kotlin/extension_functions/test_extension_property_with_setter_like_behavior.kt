// vybe-test: kotlin/extension_functions/test_extension_property_with_setter_like_behavior
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

class Holder(var value: Int)

        var Holder.doubled: Int
            get() = value * 2
            set(next) { value = next / 2 }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder(3)
            holder.doubled = 10
            __check((holder.value).toString(), "5")
            __check((holder.doubled).toString(), "10")
        }
