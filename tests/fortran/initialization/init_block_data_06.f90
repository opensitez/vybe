! vybe-test: fortran/initialization/init_block_data_06
! origin: languages/fortran/tests/fortran/test_initialization.rs
block data bd
integer::x
common /blk/ x
data x/1/
end block data bd
