;;; outlook-test.el --- tests for outlook.el         -*- lexical-binding: t; -*-

;; Copyright (C) 2018  Andrew Savonichev

;; Author: Andrew Savonichev
;; Keywords: mail

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:
;;
;; This package provides the tests for `ert'. It can be executed from
;; command line as well by running "make test"

;;; Code:
(require 'ert)
(require 'outlook)

(defun outlook-test-all ()
  (ert t))

(ert-deftest test-format-date ()
  (should (equal
           (outlook-format-date-string '(23214 50561 0))
           "Sunday, March 18, 2018 23:01")))

(ert-deftest test-html-escape ()
  (dolist (input '(("" . "")
                   ("\"" . "&quot;")
                   ("\"foo\"" . "&quot;foo&quot;")))
    (should (equal
             (outlook--html-escape-attr (car input))
             (cdr input))))

  (dolist (input '(("&<>" . "&amp;&lt;&gt;")
                   ("foo & <bar>" . "foo &amp; &lt;bar&gt;")
                   (" foo bar" . "&nbsp;foo bar")
                   ("  foo bar" . "&nbsp;&nbsp;foo bar")
                   ("   foo  bar" . "&nbsp;&nbsp; foo&nbsp;&nbsp;bar")))
    (should (equal
             (outlook--html-escape-html (car input))
             (cdr input)))))

(ert-deftest test-filter-nonstandard-tags ()
  (with-temp-buffer
    (insert "<html><body><p>foo<o:p></o:p></p></body></html>")
    (outlook--html-filter-nonstandard-tags (point-min) (point-max))
    (should (equal
             (buffer-substring (point-min) (point-max))
             "<html><body><p>foo</p></body></html>"))))

(ert-deftest test-change-charset ()
  (with-temp-buffer
    (insert "<html><head>")
    (insert "<meta http-equiv=\"Content-Type\"")
    (insert "      content=\"text/html; charset=KOI8-R\">")
    (insert "</head></html>")

    (let ((html (libxml-parse-html-region
                 (point-min) (point-max))))
      (outlook--html-change-charset html "UTF-8")

      (should (equal
               (dom-attr (car (dom-by-tag html 'meta))
                         'content)
               "text/html; charset=UTF-8")))))

(defun read-file-content (name)
  (with-temp-buffer
    (insert-file-contents name)
    (buffer-string)))

(ert-deftest test-quote-header ()
  (with-temp-buffer
    (let ((expected (read-file-content "testdata/quote-header.html")))
      (erase-buffer)
      (outlook-html-dom-print
       (outlook--html-create-quote-header
        (outlook-html-message "Alice Wheezy"
                              "Bob Squeezy <bob.squeezy@pixar.com>"
                              "Sam Potato <sam.potato@pixar.com>"
                              (outlook-format-date-string '(23214 50561 0))
                              "New Toy Story?")))
      (insert "\n")
      (should (equal expected (buffer-string))))))

(ert-deftest test-plaintext-wrap ()
  (let ((outlook-text-color-style "fancy-style")
        (expected "<p class=\"MsoNormal\"><span style=\"fancy-style\">Hello, world!</span></p>"))
    (with-temp-buffer
      (outlook-html-dom-print
       (outlook--html-wrap-plaintext-line "Hello, world!"))
      (should (equal (buffer-string) expected)))))

(ert-deftest test-plaintext-insert ()
  (with-temp-buffer
    (insert "Hey Squeeze\n\n")
    (insert "How about we start a new movie about us?\n")
    (insert "It should be fun!\n")
    (insert "\n")
    (insert "--\n")
    (insert "Alice")

    (let ((html (dom-node 'div '((class . "WordSection1"))))
          (expected (read-file-content "testdata/plaintext-wrap.html")))

      (outlook--html-insert-plaintext html (point-min) (point-max))
      (erase-buffer)
      (outlook-html-dom-print html)
      (insert "\n")

      (should (equal (buffer-string) expected)))))

(provide 'outlook-test)
;;; outlook-test.el ends here
