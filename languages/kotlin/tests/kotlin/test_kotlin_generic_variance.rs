use crate::helpers::run_prints;

#[test]
fn test_covariant_producer_can_upcast() {
    let out = run_prints(r#"
        open class Animal(val kind: String)
        class Cat(value: String) : Animal(value)
        class Dog(value: String) : Animal(value)

        class Producer<out T>(private val value: T) {
            fun value(): T = value
        }

        fun main() {
            val cats = Producer(Cat("cat"))
            val animals: Producer<Animal> = cats
            println(animals.value().kind)
            val dogs = Producer(Dog("dog"))
            val animalDogs: Producer<Animal> = dogs
            println(animalDogs.value().kind)
        }
    "#);
    assert_eq!(out, &["cat", "dog"]);
}

#[test]
fn test_contravariant_consumer_can_downcast_target_type() {
    let out = run_prints(r#"
        open class Animal(val kind: String)
        class Cat(value: String) : Animal(value)

        class Consumer<in T> {
            var value: T? = null
                private set

            fun push(v: T) {
                value = v
            }
        }

        fun main() {
            val animalConsumer = Consumer<Animal>()
            val catConsumer: Consumer<Cat> = animalConsumer
            catConsumer.push(Cat("kitty"))
            println((catConsumer.value?.kind))
        }
    "#);
    assert_eq!(out, &["kitty"]);
}
