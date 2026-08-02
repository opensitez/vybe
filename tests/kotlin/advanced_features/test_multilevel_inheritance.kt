// vybe-test: kotlin/advanced_features/test_multilevel_inheritance
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

open class Vehicle(val speed: Int)
        open class Car(speed: Int, val brand: String) : Vehicle(speed)
        class SportsCar(speed: Int, brand: String) : Car(speed, brand)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sc = SportsCar(250, "Ferrari")
            __check((sc.brand).toString(), "Ferrari")
            __check((sc.speed).toString(), "250")
        }
