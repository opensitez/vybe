! Vybe test harness — Fortran.
!
! Real Fortran: this file compiles with `gfortran` on its own. Like the COBOL
! harness it documents the shape the emitter produces rather than providing a
! subroutine to splice, because a generic `check` would need a separate
! interface per type (integer, real, logical, character) and the emitted
! comparison is a single line either way.
!
! WHY VALUES, NOT PRINTED TEXT.
! The corpus records what Vybe's logging host produced — bare, unpadded values
! like ["8", "9", "false", "true"]. gfortran's list-directed `print *` pads
! ("           8") and writes logicals as `T`/`F`, so comparing printed text
! would fail under gfortran for formatting reasons alone and the differential
! would be worthless. Comparing the VALUE is something both runtimes agree on,
! because it is semantics rather than formatting.
!
! The emitted check, replacing `print *, dst(1)` whose expected line is "8":
!
!     if ((dst(1)) /= 8) then
!         print *, "FAIL: want [8] got [", dst(1), "]"
!         stop 1
!     end if
!
! chosen by the shape of the expected text:
!
!     8        integer   ->  (x) /= 8
!     3.5      real      ->  abs((x) - 3.5) > 1.0e-6      (never compare reals with /=)
!     true     logical   ->  (x) .neqv. .true.
!     anything else      ->  trim(x) /= "..."
!
! FAILURE USES `stop <code>`, which now carries a real status: the walker used
! to lower `stop` to a bare `Return` and the profile mapped it to `noop`, so
! every `stop 1` exited 0 while gfortran gave 1. Measured after the fix:
! `stop 1` -> 1 and `error stop 2` -> 2 in both runtimes.
program vybe_check_shape
    integer :: x
    x = 8
    if ((x) /= 8) then
        print *, "FAIL: want [8] got [", x, "]"
        stop 1
    end if
    print *, "shape ok"
end program vybe_check_shape
