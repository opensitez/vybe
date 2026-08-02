// vybe-test: kotlin/enums/test_enum_array_iteration_sum
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Digit { D0, D1, D2, D3 }
fun main() { var n = 0
for (d in arrayOf(Digit.D0, Digit.D1, Digit.D2, Digit.D3)) { n += d }
println(n) }

