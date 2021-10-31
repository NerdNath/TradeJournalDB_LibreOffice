CREATE TABLE "TradeList" (
  trade_id INT GENERATED BY DEFAULT AS IDENTITY,
  ticker VARCHAR(6),
  exchange VARCHAR(6),
  trade_type VARCHAR(5),
  approx_float BIGINT,
  pattern_name VARCHAR(40),
  q1 VARCHAR(3),
  q2 VARCHAR(3),
  q3 VARCHAR(3),
  q4 VARCHAR(3),
  SSS_PP INT,
  SSS_RR INT,
  SSS_EE INT,
  SSS_PF INT,
  SSS_TS INT,
  SSS_RC INT,
  SSS_ME INT,
  PRIMARY KEY(trade_id)
);