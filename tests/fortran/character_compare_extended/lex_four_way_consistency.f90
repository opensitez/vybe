! vybe-test: fortran/character_compare_extended/lex_four_way_consistency
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('p','q')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('p','q'), "]"
    stop 1
end if
if ((lgt('q','p')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('q','p'), "]"
    stop 1
end if
if ((lle('p','q')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle('p','q'), "]"
    stop 1
end if
if ((lge('q','p')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('q','p'), "]"
    stop 1
end if
end program t
