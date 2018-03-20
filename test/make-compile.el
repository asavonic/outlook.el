(setq files '("outlook.el"
              "outlook-mu4e.el"))
(add-to-list 'load-path default-directory)
(setq byte-compile--use-old-handlers nil)
(mapc #'byte-compile-file files)
