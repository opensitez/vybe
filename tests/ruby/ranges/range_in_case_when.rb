# vybe-test: ruby/ranges/range_in_case_when
# origin: languages/ruby/tests/ruby/test_ranges.rs
# vybe-test-mode: compile

score = 75
case score
when 90..100 then puts 'A'
when 70..89  then puts 'B'
else              puts 'C'
end
