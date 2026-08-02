// vybe-test: kotlin/properties/test_property_nullable_value_default_null
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Note {
            var text: String? = null
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val note = Note()
            __check((note.text == null).toString(), "true")
            note.text = "ok"
            __check((note.text).toString(), "ok")
        }
