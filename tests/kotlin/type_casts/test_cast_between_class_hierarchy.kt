// vybe-test: kotlin/type_casts/test_cast_between_class_hierarchy
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

open class Vehicle(val speed: Int)
        class Car(speed: Int) : Vehicle(speed)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Vehicle = Car(120)
            val car = value as Car
            __check((car.speed).toString(), "120")
        }
