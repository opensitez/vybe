-- vybe-test: lua/operators/logical_and_short_circuit_shape
-- origin: languages/lua/tests/lua/test_operators.rs

local x = false and print("skip")
