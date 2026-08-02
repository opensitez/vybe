// vybe-test: kotlin/properties/test_property_same_name_local_and_member_do_not_interfere
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder {
            val value = "member"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "local"
            val holder = Holder()
            __check((value).toString(), "local")
            __check((holder.value).toString(), "member")
        }
