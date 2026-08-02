// vybe-test: kotlin/advanced_features/test_advanced_override_open_property
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

open class Vehicle {
            open val kind: String = "vehicle"
        }

        class Car : Vehicle() {
            override val kind: String = "car"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v: Vehicle = Car()
            __check((v.kind).toString(), "car")
        }
