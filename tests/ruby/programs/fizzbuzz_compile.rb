# vybe-test: ruby/programs/fizzbuzz_compile
# origin: languages/ruby/tests/ruby/test_programs.rs
# vybe-test-mode: compile


for i in 1..15
  if i % 15 == 0
    puts "FizzBuzz"
  elsif i % 3 == 0
    puts "Fizz"
  elsif i % 5 == 0
    puts "Buzz"
  else
    puts i
  end
end
