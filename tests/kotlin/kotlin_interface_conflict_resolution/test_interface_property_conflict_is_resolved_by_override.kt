// vybe-test: kotlin/kotlin_interface_conflict_resolution/test_interface_property_conflict_is_resolved_by_override
// origin: languages/kotlin/tests/kotlin/test_kotlin_interface_conflict_resolution.rs

interface Marker {
            val label: String
                get() = "marker"
        }

        interface Debug {
            val label: String
                get() = "debug"
        }

        class Item : Marker, Debug {
            override val label: String = "item"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Item().label).toString(), "item")
        }
