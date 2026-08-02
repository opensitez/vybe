// vybe-test: kotlin/advanced_features/test_class_inheritance_and_override
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

open class Animal(val name: String) {
            open fun speak() {
                println(name + " makes a sound")
            }
        }

        class Dog(name: String) : Animal(name) {
            override fun speak() {
                println(name + " barks")
            }
        }

        fun main() {
            val dog = Dog("Rex")
            dog.speak()
        }

