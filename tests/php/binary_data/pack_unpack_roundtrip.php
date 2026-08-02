<?php
// vybe-test: php/binary_data/pack_unpack_roundtrip
// origin: languages/php/tests/php/test_binary_data.rs
// vybe-test-mode: compile

$data = ['x' => 100, 'y' => 200, 'z' => 300];
$packed = pack('NNN', $data['x'], $data['y'], $data['z']);
$out = unpack('Nx/Ny/Nz', $packed);
echo $out['x'] . ',' . $out['y'] . ',' . $out['z'];
