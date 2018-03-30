;;; outlook.el --- send emails in MS Outlook style  -*- lexical-binding: t; -*-

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

;; This package provides routines to add a plain-text reply into HTML
;; formatted email using MS Outlook style of quotation. Integrations
;; into mail packages (mu4e, gnus) should be provided by separate
;; outlook-xxx.el files.

;;; Code:
(require 'dom)
(require 'seq)
(require 'shr)

(defcustom outlook-reply-separator-style
  "border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm 0cm 0cm"
  "CSS style to apply for the first quoted div"
  :type 'string :group 'outlook)

(defcustom outlook-text-color-style
  "color:#1F497D;mso-fareast-language:EN-US"
  "CSS style to apply for the reply text."
  :type 'string :group 'outlook)

(defcustom outlook-organization-domain-regexp nil
  "Domain name of your current organization. If non-nil, all
   recipient email addresses will be checked for belonging to
   this domain and you will be queried for confirmation
   otherwise."
  :type 'regexp :group 'outlook)

(defun outlook-html-read (start end)
  "Read HTML region into a DOM structure and prepare it for a
reply insertion."
  (let ((orig (current-buffer)))
    (with-temp-buffer
      (insert-buffer-substring orig start end)
      (outlook--html-filter-nonstandard-tags
       (point-min) (point-max))

      (let ((html (libxml-parse-html-region
                   (point-min) (point-max))))
        (outlook--html-change-charset html "UTF-8")
        html))))

(defun outlook-message (from to cc date subject)
  "Create a parent message descriptor. Any arguments can be nil,
in which case the corresponding field will be missing in
quotation header."
  `((from    . ,from)
    (to      . ,to)
    (cc      . ,cc)
    (date    . ,date)
    (subject . ,subject)))

(defun outlook-html-insert-reply (html message start end)
  "Inserts a REPLY block into HTML email."

  (unless (eq (dom-tag html) 'html)
    (error "Not an html DOM"))

  (outlook--html-insert-reply-quote html message)
  (outlook--html-insert-plaintext html start end))

(defvar outlook--html-self-closing-tags
  '(area base br col command embed hr img
         input keygen link menuitem meta
         param source track wbr))

;; FIXME: this is actually a copy of shr-dom-print, but it also
;; properly handles self-closed tags, while the original functions is
;; not aware of them. Another change is quote-escape for attribute
;; strings and xml-escape for string nodes. This wasn't done in
;; shr-dom-print for some reason. It might be better to contribute
;; these changes back.
(defun outlook-html-dom-print (dom)
  "Convert DOM into a string containing the xml representation."
  (insert (format "<%s" (dom-tag dom)))
  (dolist (attr (dom-attributes dom))
    ;; Ignore attributes that start with a colon because they are
    ;; private elements.
    (unless (= (aref (format "%s" (car attr)) 0) ?:)
      (insert (format " %s=\"%s\""
                      (car attr)
                      (outlook--html-escape-attr (cdr attr))))))

  ;; Skip if node is self-closing.
  (if (and (member (dom-tag dom) outlook--html-self-closing-tags)
           (eq nil (dom-children dom)))
      (insert " />")

    ;; Not a self closing tag: call recursively on children.
    (insert ">")
    (let (url)
      (dolist (elem (dom-children dom))
        (cond
         ((stringp elem)
          (insert (outlook--html-escape-html elem)))
         ((eq (dom-tag elem) 'comment)
          )
         ((or (not (eq (dom-tag elem) 'image))
              ;; Filter out blocked elements inside the SVG image.
              (not (setq url (dom-attr elem ':xlink:href)))
              (not shr-blocked-images)
              (not (string-match shr-blocked-images url)))
          (outlook-html-dom-print elem)))))
    (insert (format "</%s>" (dom-tag dom)))))

(defun outlook-format-date-string (date)
  "Format date-time structure following the outlook rules."
  (format-time-string "%A, %B %d, %G %H:%M" date))

(defun outlook-format-sender (name email)
  "Returns a string to put into From: field of a quotation
header. Usually it is just a NAME, but EMAIL is returned if NAME
is nil."
  (if name name email))

(defun outlook-format-recipient (name email)
  "Returns a string to put into To: or Cc: fields of a quotation
header. Return format is 'Name Lastname <foo@email.com>' if NAME
is non-nil, otherwise EMAIL is returned."
  (if name
      (format "%s <%s>" name email)
    email))

(defun outlook--html-wrap-plaintext-line (line)
  "Wrap plaintext region into ugly Microsoft-specific html tags,
  e.g. <p class=MsoNormal>."

  (if (string= "" line)
      (dom-node 'br)
    (dom-node
     'p '((class . "MsoNormal"))
     (dom-node
      'span `((style . ,outlook-text-color-style))
      line))))

(defun outlook--html-insert-plaintext (html start end)
  "Insert a plaintext block into DEST-NODE, preserving linebreaks."
  (let ((dest (outlook--html-find-insert-pt html))
        lines)
    (dom-add-child-before dest (dom-node 'br))
    (save-excursion
      (goto-char start)
      (while (< (point) end)
        (push (outlook--html-wrap-plaintext-line
               (buffer-substring-no-properties
                (point) (min (line-end-position) end)))
              lines)
        (forward-line 1))

      (dolist (line lines dest)
        (dom-add-child-before dest line)))))

(defun outlook--html-remove-reply-separator (html)
  (let ((sep-tag (seq-find
                  (lambda (a) (equal (dom-attr a 'name)
                                     "_____replyseparator"))
                  (dom-by-tag html 'a))))
    (delq sep-tag (dom-parent html sep-tag))))

(defun outlook--html-find-insert-pt (html)
  (or (car (dom-by-class html "WordSection1"))
      (dom-child-by-tag (dom-child-by-tag html 'body)
                        'div)))

(defun outlook--html-insert-reply-quote (html message)
  (let ((div (outlook--html-find-insert-pt html)))
    (outlook--html-remove-reply-separator html)
    (dom-add-child-before div (dom-node 'br))
    (dom-add-child-before
     div (outlook--html-create-quote-header message))))

(defun outlook--html-insert-quote-header-field (node field-name value)
  (when value
    (dom-append-child node (dom-node 'br))
    (dom-append-child node (dom-node 'b nil field-name))
    (dom-append-child node (concat " " value))))

(defun outlook--html-create-quote-header (message)
  (let* ((div (dom-node
               'div `((style . ,outlook-reply-separator-style))))
         (par (dom-node 'p '((class . "MsoNormal")))))

    (dom-append-child div par)

    (dom-append-child
     par (dom-node 'a '((name . "_____replyseparator"))))

    (dom-append-child
     par (dom-node
          'b nil (dom-node
                  'span '((lang . "EN-US")) "From: ")))

    (let ((span (dom-node 'span '((lang . "EN-US")))))
      (dom-append-child par span)

      (dom-append-child span (cdr (assq 'from message)))

      (outlook--html-insert-quote-header-field
       span "Sent:" (cdr (assq 'date message)))

      (outlook--html-insert-quote-header-field
       span "To:" (cdr (assq 'to message)))

      (outlook--html-insert-quote-header-field
       span "Cc:" (cdr (assq 'cc message)))

      (outlook--html-insert-quote-header-field
       span "Subject:" (cdr (assq 'subject message))))

    ;; Insert visual separator b/w quote and the actual email
    (dom-set-attribute div 'style outlook-reply-separator-style)

    ;; Wrap it into yet another div (who knows why?)
    (dom-node 'div nil div)))


(defun outlook--html-filter-nonstandard-tags (start end)
  (save-excursion
    (goto-char start)
    (while (re-search-forward "<o:p>\\(.*\\)</o:p>" end t)
      (replace-match "\\1"))))


(defun outlook--html-change-charset (html desired-charset)
  (let ((meta
         (seq-find (lambda (elem) (equal "Content-Type"
                                         (dom-attr elem 'http-equiv)))
                   (dom-by-tag html 'meta))))
    (dom-set-attribute
     meta 'content
     (replace-regexp-in-string
      "charset=\\S-+" (format "charset=%s" desired-charset)
      (dom-attr meta 'content)))))

(defvar outlook--html-escape-html-replacements
  '(("&"   . "&amp;")
    ("<"   . "&lt;")
    (">"   . "&gt;")
    ("^ "  . "&nbsp;")
    ("\\(&nbsp;\\| \\) " . "&nbsp;&nbsp;")))

(defun outlook--html-escape-html (str)
  (with-temp-buffer
    (insert str)
    (dolist (regex-rep outlook--html-escape-html-replacements
                       (buffer-string))
      (goto-char (point-min))
      (while (re-search-forward (car regex-rep) nil t)
        (replace-match (cdr regex-rep))))))

(defun outlook--html-escape-attr (str)
  (replace-regexp-in-string "\"" "&quot;" str t 'literal))

(defun outlook--txt-insert-quote-header-field (field-name value)
  (when value
    (insert (format "%s %s\n" field-name value))))

(defun outlook-txt-insert-quote-header (message)
  (insert "-----Original Message-----\n")

  (outlook--txt-insert-quote-header-field
   "From:" (cdr (assq 'from message)))

  (outlook--txt-insert-quote-header-field
   "Sent:" (cdr (assq 'date message)))

  (outlook--txt-insert-quote-header-field
   "To:" (cdr (assq 'to message)))

  (outlook--txt-insert-quote-header-field
   "Cc:" (cdr (assq 'cc message)))

  (outlook--txt-insert-quote-header-field
   "Subject:" (cdr (assq 'subject message))))

(defun outlook--recipients-outside-organization (recipients)
  "Find addresses among RECIPIENTS that do not belong to the organization domain.
  Should be adviced for `message-bogus-recipient-p'."

  (when outlook-organization-domain-regexp
    (let (found)
      (mapc (lambda (address)
              (setq address (or (cadr address) ""))
              (unless (string-match-p
                       (concat ".*@" outlook-organization-domain-regexp "$")
                       address)
                (push address found)))
            (mail-extract-address-components recipients t))
      found)))

(eval-after-load 'message
  (advice-add 'message-bogus-recipient-p :before-until
              #'outlook--recipients-outside-organization))

(provide 'outlook)
;;; outlook.el ends here

