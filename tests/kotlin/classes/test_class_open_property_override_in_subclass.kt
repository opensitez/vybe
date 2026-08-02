// vybe-test: kotlin/classes/test_class_open_property_override_in_subclass
// origin: languages/kotlin/tests/kotlin/test_classes.rs

open class Vehicle {
            open val tag: String = "vehicle"
        }

        class Truck : Vehicle() {
            override val tag: String = "truck"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v: Vehicle = Truck()
            __check((v.tag).toString(), "truck")
        }
