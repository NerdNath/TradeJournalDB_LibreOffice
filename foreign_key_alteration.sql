
ALTER TABLE "Orders" 
ADD FOREIGN KEY(trade_id) REFERENCES "TradeList"(trade_id);

ALTER TABLE "Executions"
ADD FOREIGN KEY(account_number,order_number) 
REFERENCES "Orders"(account_number,order_number);

ALTER TABLE "Executions"
ADD FOREIGN KEY(trade_id)
REFERENCES "TradeList"(trade_id);