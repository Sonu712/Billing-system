SQL>CREATE USER PROJECT IDENTIFIED BY SONU
SQL>GRANT  CONNECT,RESOURCE,DBA  TO  SONU 
SQL:- CONN USER-NAME: PROJECT/SONU
SQL> CREATE TABLE FERTDETAIL
  (FERTNM VARCHAR2(20) ,
  FERTCOD VARCHAR2(20) PRIMARY KEY,
  COMPNYNM VARCHAR2(20),
  BCHNO VARCHAR2(20),
  MFGDT DATE,
  EXPDT DATE,
  QTYBAG NUMBER(8,2),
  PCEBAG NUMBER(8,2),
  PCEKG NUMBER(8,2));
SQL> CREATE TABLE PESTDETAIL
  (PESTNM VARCHAR2(20) ,
  PESTCOD VARCHAR2(20) PRIMARY KEY,
  COMPNYNM VARCHAR2(20),
  BCHNO VARCHAR2(20),
  MFGDT DATE,
  EXPDT DATE,
  QTYBAG NUMBER(8,2),
  PCEBAG NUMBER(8,2),
  PCEKG NUMBER(8,2));
SQL> CREATE TABLE SEEDDETAIL
  (SEEDNM VARCHAR2(20) ,
  SEEDCOD VARCHAR2(20) PRIMARY KEY,
  COMPNYNM VARCHAR2(20),
  BCHNO VARCHAR2(20),
  MFGDT DATE,
  EXPDT DATE,
  QTYBAG NUMBER(8,2),
  PCEBAG NUMBER(8,2),
  PCEKG NUMBER(8,2));
SQL>CREATE TABLE SUPPDETAIL
    (SUPP_ID VARCHAR2(20) PRIMARY KEY ,
    SUPP_NM VARCHAR2(20),
    SUPP_ADD VARCHAR2(50),
    STATE VARCHAR2(20),
    MOB NUMBER(8,2),
    AC_NO NUMBER (8,2),
    IFSC VARCHAR2(15))
SQL>CREATE TABLE CUSTDETAIL
    (ADNO NUMBER(12,2) PRIMARY KEY,
     NM VARCHAR2(20),
     ADS VARCHAR2(50),
     PIN NUMBER(8,2),
     MBNO NUMBER(8,2))
SQL>CREATE TABLE SALE
( INVNO    VARCHAR2(10) PRIMARY KEY ,
      ADHAR    NUMBER(12),
     TTL      NUMBER(8,2),
     DATE1    DATE)
SQL>CREATE TABLE PURCHASE
( ORDR   VARCHAR2(10) PRIMARY KEY ,
  TTL      NUMBER(8,2),
  DATE1    DATE,
  sid varchar2(20))
  SQL> CREATE TABLE LOGIN 
(USER_ID        VARCHAR2(20)
 PASS        VARCHAR2(20)
 MOB                  NUMBER(10))
SQL>create table order_item (
SNO      NUMBER(6),
PRD      VARCHAR2(150),
QTY      NUMBER(8),
RATE     NUMBER(8,2),
AMT      NUMBER(8,2),
ORDR VARCHAR2(10),
FOREIGN KEY (ORDR) REFERENCES PURCHASE(ORDR) );
SQL>create table sale_item (
SNO      NUMBER(6),
PRD      VARCHAR2(150),
QTY      NUMBER(8),
RATE     NUMBER(8,2),
AMT      NUMBER(8,2),
INVNO    VARCHAR2(10),
FOREIGN KEY (INVNO) REFERENCES SALE(INVNO) );