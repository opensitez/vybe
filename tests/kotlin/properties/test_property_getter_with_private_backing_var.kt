// vybe-test: kotlin/properties/test_property_getter_with_private_backing_var
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Meter {
            private var raw: Int = 2
            val doubled: Int
                get() = raw * 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Meter().doubled).toString(), "4")
        }
