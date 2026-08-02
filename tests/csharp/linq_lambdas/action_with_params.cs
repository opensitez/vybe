// vybe-test: csharp/linq_lambdas/action_with_params
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

Action<string, int> describe = (name, age) => {
    Console.WriteLine(name + " is " + age);
};
describe("Alice", 30);
describe("Bob", 25);
