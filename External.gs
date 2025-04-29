{\rtf1\ansi\ansicpg1252\cocoartf2761
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red20\green67\blue174;\red246\green247\blue249;\red46\green49\blue51;
\red24\green25\blue27;\red186\green6\blue115;\red162\green0\blue16;\red77\green80\blue85;}
{\*\expandedcolortbl;;\cssrgb\c9412\c35294\c73725;\cssrgb\c97255\c97647\c98039;\cssrgb\c23529\c25098\c26275;
\cssrgb\c12549\c12941\c14118;\cssrgb\c78824\c15294\c52549;\cssrgb\c70196\c7843\c7059;\cssrgb\c37255\c38824\c40784;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs26 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 _getData\cf4 \strokec4 (\cf5 \strokec5 params\cf4 \strokec4 )\{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf5 \strokec5 params\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 doc\cf4 \strokec4  = \cf5 \strokec5 params\cf4 \strokec4 .\cf6 \strokec6 SpreadsheetApp\cf4 \strokec4 .\cf5 \strokec5 openById\cf4 \strokec4 (\cf5 \strokec5 params\cf4 \strokec4 .\cf6 \strokec6 SCRIPT_PROP\cf4 \strokec4 .\cf5 \strokec5 getProperty\cf4 \strokec4 (\cf7 \cb3 \strokec7 "key"\cf4 \cb3 \strokec4 ));\cb1 \
\cb3   \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 sheet\cf4 \strokec4  = \cf5 \strokec5 doc\cf4 \strokec4 .\cf5 \strokec5 getSheetByName\cf4 \strokec4 (\cf5 \strokec5 params\cf4 \strokec4 .\cf5 \strokec5 sheet_name\cf4 \strokec4 );\cb1 \
\cb3   \cf5 \strokec5 values\cf4 \strokec4  = \cf5 \strokec5 sheet\cf4 \strokec4 .\cf5 \strokec5 getDataRange\cf4 \strokec4 ().\cf5 \strokec5 getValues\cf4 \strokec4 ();\cb1 \
\cb3   \cf2 \strokec2 return\cf4 \strokec4  \cf5 \strokec5 valuesToObject\cf4 \strokec4 (\cf5 \strokec5 values\cf4 \strokec4 );\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 _sendMail\cf4 \strokec4 (\cf5 \strokec5 data\cf4 \strokec4 , \cf6 \strokec6 MAIL_APP\cf4 \strokec4 )\{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cb1 \
\cb3     \cf2 \strokec2 if\cf4 \strokec4 (!\cf5 \strokec5 data\cf4 \strokec4 .\cf6 \strokec6 Email\cf4 \strokec4 )\cb1 \
\cb3     \cf2 \strokec2 return\cf4 \strokec4  \{\cf5 \strokec5 success\cf4 \strokec4 : \cf2 \strokec2 false\cf4 \strokec4 , \cf5 \strokec5 msg\cf4 \strokec4 : \cf7 \cb3 \strokec7 'Kh\'f4ng t\'ecm th\uc0\u7845 y Email kh\'e1ch h\'e0ng'\cf4 \cb3 \strokec4 \}\cb1 \
\
\cb3     \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 emailAddress\cf4 \strokec4  = \cf5 \strokec5 data\cf4 \strokec4 .\cf6 \strokec6 Email\cf4 \strokec4 ; \cf8 \cb3 \strokec8 // First column\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 content\cf4 \strokec4  = \cf5 \strokec5 data\cf4 \strokec4 .\cf5 \strokec5 content\cf4 \strokec4 ; \cf8 \cb3 \strokec8 // Second column\cf4 \cb1 \strokec4 \
\cb3     \cf2 \strokec2 var\cf4 \strokec4  \cf5 \strokec5 subject\cf4 \strokec4  = \cf5 \strokec5 data\cf4 \strokec4 .\cf5 \strokec5 subject\cf4 \strokec4 ;\cb1 \
\cb3     \cf6 \strokec6 MAIL_APP\cf4 \strokec4 .\cf5 \strokec5 sendEmail\cf4 \strokec4 (\cf5 \strokec5 emailAddress\cf4 \strokec4 , \cf5 \strokec5 subject\cf4 \strokec4 , \cf7 \cb3 \strokec7 ''\cf4 \cb3 \strokec4 , \{\cf5 \strokec5 htmlBody\cf4 \strokec4 : \cf5 \strokec5 content\cf4 \strokec4 \});\cb1 \
\cb3     \cf2 \strokec2 return\cf4 \strokec4  \{\cf5 \strokec5 data\cf4 \strokec4 , \cf5 \strokec5 quota\cf4 \strokec4 : \cf6 \strokec6 MAIL_APP\cf4 \strokec4 .\cf5 \strokec5 getRemainingDailyQuota\cf4 \strokec4 (), \cf5 \strokec5 success\cf4 \strokec4 : \cf2 \strokec2 true\cf4 \strokec4 \}\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 _getQuota\cf4 \strokec4 (\cf6 \strokec6 MAIL_APP\cf4 \strokec4 )\{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf6 \strokec6 Logger\cf4 \strokec4 .\cf5 \strokec5 log\cf4 \strokec4 (\cf6 \strokec6 MAIL_APP\cf4 \strokec4 .\cf5 \strokec5 getRemainingDailyQuota\cf4 \strokec4 ())\cb1 \
\cb3   \cf2 \strokec2 return\cf4 \strokec4  \{\cf5 \strokec5 value\cf4 \strokec4 : \cf6 \strokec6 MAIL_APP\cf4 \strokec4 .\cf5 \strokec5 getRemainingDailyQuota\cf4 \strokec4 ()\};\cb1 \
\cb3 \}\cb1 \
}
