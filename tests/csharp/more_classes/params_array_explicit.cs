// vybe-test: csharp/more_classes/params_array_explicit
// origin: languages/csharp/tests/csharp/test_more_classes.rs

class Program {
            static int Sum(params int[] numbers) {
                var total = 0;
                for (var i = 0; i < 5; i++) {
                    total = total + numbers[i];
                }
                return total;
            }
        }
        var arr = new int[] {1, 2, 3, 4, 5};
        Console.WriteLine(Program.Sum(arr));
