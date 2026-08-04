# vybe-test: ruby/symbols/symbol_in_case_when
# origin: languages/ruby/tests/ruby/test_symbols.rs
# vybe-test-mode: compile

status = :ok
result = case status
when :ok then 'good'
when :error then 'bad'
else 'unknown'
end
