! vybe-test: fortran/allocation/allocate_character_payload_after_alloc_then_assign
! origin: languages/fortran/tests/fortran/test_allocation.rs
program t
character(len=:), allocatable :: s
allocate(character(len=4) :: s)
s = 'rust'
if ((len(s)) /= 4) then
    print *, "FAIL: want [4] got [", len(s), "]"
    stop 1
end if
if (trim(s) /= "rust") then
    print *, "FAIL: want [rust] got [", s, "]"
    stop 1
end if
end program t
