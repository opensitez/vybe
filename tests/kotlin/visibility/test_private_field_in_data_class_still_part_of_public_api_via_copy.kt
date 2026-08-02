// vybe-test: kotlin/visibility/test_private_field_in_data_class_still_part_of_public_api_via_copy
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

data class Holder(private val secret: String, val value: String) {
            fun reveal(): String = secret
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val holder = Holder("x", "y")
            __check((holder.value).toString(), "y")
            __check((holder.reveal()).toString(), "x")
            __check((holder.copy(value = "z").value).toString(), "z")
        }
