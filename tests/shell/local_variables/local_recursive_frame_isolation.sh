#!/usr/bin/env bash
# vybe-test: bash/local_variables/local_recursive_frame_isolation
# Each recursive function invocation instantiates its own isolated local variable storage frame.
fail() { printf 'FAIL: %s\n' "$*"; exit 1; }
recurse() {
    local depth=$1
    local frame_id="depth_${depth}"
    if [ "$depth" -gt 1 ]; then
        recurse $(( depth - 1 ))
    fi
    # After recursive call returns, frame_id must still equal this frame's value
    [ "$frame_id" = "depth_${depth}" ] || fail "frame corrupted at depth $depth: got [$frame_id]"
}
recurse 4
echo PASS
exit 0
