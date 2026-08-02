// vybe-test: kotlin/properties/test_property_primary_constructor_mutable_property
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Box(var item: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box("start")
            box.item = "done"
            __check((box.item).toString(), "done")
        }
