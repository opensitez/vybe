! vybe-test: fortran/move_alloc/move_alloc_10
! origin: languages/fortran/tests/fortran/test_move_alloc.rs
program p
class(*), allocatable :: a,b
allocate(integer :: a)
call move_alloc(a,b)
end program p
