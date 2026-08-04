-- vybe-test: lua/loops_for_generic/test_for_gen_pairs
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local s=''; local t={a=1,b=2}; for k,v in pairs(t) do s=s..k..v end; -- Can't strictly assert string order due to hash, so just count
         local count=0; for _ in pairs(t) do count=count+1 end; print(count)
