#!/usr/bin/env bash
set -euo pipefail

# Run each Python test under the vybe runner `vybec` and require each to print PASS
ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
VYBEC="$ROOT_DIR/target/debug/vybec"
TEST_DIR="$ROOT_DIR/examples/python/tests"

if [ ! -x "$VYBEC" ]; then
  echo "vybec not built. Run: cargo build --bin vybec"
  exit 2
fi

failed=0
for t in "$TEST_DIR"/*.py; do
  echo "Running $(basename "$t")..."
  out="$($VYBEC "$t" 2>&1)"
  echo "$out"
  if echo "$out" | grep -q "PASS"; then
    echo "  OK"
  else
    echo "  FAIL"
    failed=$((failed+1))
  fi
done

if [ $failed -ne 0 ]; then
  echo "$failed test(s) failed"
  exit 1
fi

echo "All Python tests passed"
