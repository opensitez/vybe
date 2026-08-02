// vybe-test: kotlin/interfaces/test_interface_in_while_condition
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Marker { fun hit(): Boolean }
class Yes: Marker { override fun hit(): Boolean = true }
class No: Marker { override fun hit(): Boolean = false }
fun main() { var score = 0
for (m in arrayOf(Yes(), No())) { if (m.hit()) score += 1 }
println(score) }

