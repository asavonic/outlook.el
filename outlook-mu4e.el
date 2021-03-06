;;; outlook-mu4e.el --- integration with `mu4e' package  -*- lexical-binding: t; -*-

;; Copyright (C) 2018  Andrew Savonichev

;; Author: Andrew Savonichev
;; URL: https://github.com/asavonic/outlook.el
;; Version: 0.1
;; Keywords: mail

;; Package-Requires: ((emacs "24.4"))

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

;; This is an integration package to use outlook.el with `mu4e' MUA
;; package.
;;
;; Use `outlook-mu4e-html-message-finalize' before sending a reply in
;; order to format it. Run `outlook-mu4e-html-message-preview' to
;; preview the final message before sending it.

;;; Code:
(require 'outlook)
(require 'mu4e)
(require 'subr-x)

(defun outlook-mu4e-message-finalize ()
  (interactive)
  (let ((message (outlook-mu4e-parent-message))
        (html-body (outlook-mu4e-parent-html-body))
        (txt-body (outlook-mu4e-parent-txt-body)))

    (cond
     (txt-body
      (outlook-mu4e-txt-message-finalize message txt-body))
     (html-body
      (outlook-mu4e-html-message-finalize message html-body))
     (t (error "Cannot find parent message body.")))))

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

    (when (outlook-mu4e-parent-html-body)
      (unless (search-forward "<#part type=\"text/html\">" nil t)
        (error "Did you call outlook-mu4e-html-message-finalize before?")))

    (let ((temp-file (make-temp-file "emacs-email")))
      (write-region (point) (point-max) temp-file)
      (browse-url (concat "file://" temp-file)))))

(defun outlook-mu4e-parent-html-body ()
  (let ((html-string
         (plist-get mu4e-compose-parent-message :body-html)))
    (when html-string
      (with-temp-buffer
        (insert html-string)
        (outlook-html-read (point-min) (point-max))))))

(defun outlook-mu4e-parent-message ()
  (outlook-message
   (outlook-mu4e-format-sender
    (car (plist-get mu4e-compose-parent-message
                    :from)))

   (outlook-mu4e-format-recipients-list
    (plist-get mu4e-compose-parent-message
               :to))

   (outlook-mu4e-format-recipients-list
    (plist-get mu4e-compose-parent-message
               :cc))

   (outlook-format-date-string
    (plist-get mu4e-compose-parent-message
               :date))

   (plist-get mu4e-compose-parent-message
              :subject)))

(defun outlook-mu4e-format-sender (contact)
  (when contact
    (outlook-format-sender (car contact) (cdr contact))))

(defun outlook-mu4e-format-recipients-list (contacts)
  (when contacts
    (string-join
     (mapcar (lambda (name-email)
               (outlook-format-recipient (car name-email) (cdr name-email)))
             contacts)
     "; ")))


(defun outlook-mu4e-parent-txt-body ()
  (plist-get mu4e-compose-parent-message :body-txt))

(defun outlook-mu4e-txt-message-finalize (message txt)
  (save-excursion
    (goto-char (point-max))
    (insert "\n\n")
    (outlook-txt-insert-quote-header message)
    (insert "\n")
    (insert txt)))

(provide 'outlook-mu4e)
;;; outlook-mu4e.el ends here

