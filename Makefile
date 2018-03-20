MU4E ?= /usr/share/emacs/site-lisp/mu4e/

ADD_MU4E = --eval "(add-to-list 'load-path \"$(MU4E)\")"

compile: clean
	emacs -batch $(ADD_MU4E) -l test/make-compile.el

clean:
	rm -f *.elc

test:
	emacs  -batch   $(ADD_MU4E)             \
			-l outlook.el           \
			-l outlook-mu4e.el      \
			-l outlook-test.el      \
			-e outlook-test-all

.PHONY: update compile clean test
