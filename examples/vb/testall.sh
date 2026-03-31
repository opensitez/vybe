#!/usr/bin/env bash
for f in tests/test_*.vb 
do 
  result=$(cargo run --quiet --bin vybe -- "$f" 2>&1 | tail -1)
  echo "$(basename $f): $result"
done

