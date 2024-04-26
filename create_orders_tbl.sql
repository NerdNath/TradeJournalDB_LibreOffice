
CREATE TABLE "Orders" (
  trade_id INT,
  account_number INT,
  order_number BIGINT,
  action VARCHAR(4),
  submit_date DATE,
  exe_type VARCHAR(3),
  requested_price NUMERIC(10,5),
  status VARCHAR(20),
  qty_requested INT,
  PRIMARY KEY(account_number, order_number),
  CONSTRAINT UC_order UNIQUE(trade_id,order_number)
);