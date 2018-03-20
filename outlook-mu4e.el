;;; outlook-mu4e.el --- integration with `mu4e' package  -*- lexical-binding: t; -*-

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

;;; Code:
(require 'outlook)
(require 'mu4e)
(require 'subr-x)

(defun outlook-mu4e-message-finalize ()
  (interactive)
  (let ((message (outlook-mu4e-parent-message))
        (html-body (outlook-mu4e-parent-html-body)))

    (if html-body
        (outlook-mu4e-html-message-finalize message html-body)
      (error "Plaintext message is not supported in outlook.el yet."))))

(defun outlook-mu4e-html-message-finalize (message html)
  (message-goto-body)
  (outlook-html-insert-reply html message (point) (point-max))
  (delete-region (point) (point-max))
  (insert "<#part type=\"text/html\">\n")
  (outlook-html-dom-print html))

(defun outlook-mu4e-html-message-preview ()
  (interactive)
  (save-excursion
    (message-goto-body)
    (forward-line)
    (let ((temp-file (make-temp-file "emacs-email")))
      (write-region (point) (point-max) temp-file)
      (browse-url (concat "file://" temp-file)))))

(defun outlook-mu4e-parent-html-body ()
  (let ((html-string
         (plist-get mu4e-compose-parent-message :body-html)))
    (with-temp-buffer
      (insert html-string)
      (outlook-html-read (point-min) (point-max)))))

(defun outlook-mu4e-parent-message ()
  (outlook-html-message
   (outlook-mu4e-format-contacts-list-no-email
    (plist-get mu4e-compose-parent-message
               :from))

   (outlook-mu4e-format-contacts-list
    (plist-get mu4e-compose-parent-message
               :to))

   (outlook-mu4e-format-contacts-list
    (plist-get mu4e-compose-parent-message
               :cc))

   (outlook-format-date-string
    (plist-get mu4e-compose-parent-message
               :date))

   (plist-get mu4e-compose-parent-message
              :subject)))

(defun outlook-mu4e-format-contacts-list-no-email (contacts)
  (when contacts
    (string-join
     (mapcar (lambda (name-email) (format "%s" (car name-email)))
             contacts)
     "; ")))

(defun outlook-mu4e-format-contacts-list (contacts)
  (when contacts
    (string-join
     (mapcar (lambda (name-email)
               (format "%s <%s>" (car name-email) (cdr name-email)))
             contacts)
     "; ")))

(provide 'outlook-mu4e)
;;; outlook-mu4e.el ends here

