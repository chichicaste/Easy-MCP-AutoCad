from sqlalchemy import create_engine, MetaData, Table, Column, inspect
from sqlalchemy.exc import SQLAlchemyError


class DatabaseManager:
    def __init__(self, connection_string):
        """Initialize the database connection manager"""
        self.connection_string = connection_string
        self.engine = None
        self.metadata = None
        self.inspector = None

    def connect(self):
        """Connect to the database"""
        try:
            self.engine = create_engine(self.connection_string)
            self.metadata = MetaData()
            self.metadata.reflect(bind=self.engine)
            self.inspector = inspect(self.engine)
            return True
        except Exception as e:
            print(f"Failed to connect to the database: {str(e)}")
            return False

    def disconnect(self):
        """Disconnect from the database"""
        if self.engine:
            self.engine.dispose()

    def get_all_tables(self):
        """Get all table names"""
        try:
            return self.inspector.get_table_names()
        except SQLAlchemyError as e:
            return f"Failed to retrieve table list: {str(e)}"

    def get_table_schema(self, table_name):
        """Get schema information for a specific table"""
        try:
            if table_name not in self.metadata.tables:
                return f"Table '{table_name}' does not exist"

            columns = []
            for column in self.inspector.get_columns(table_name):
                columns.append({
                    "name": column['name'],
                    "type": str(column['type']),
                    "nullable": column['nullable'],
                    "default": str(column['default']) if column['default'] else None
                })

            # Primary key
            primary_keys = self.inspector.get_pk_constraint(table_name)

            # Foreign keys
            foreign_keys = []
            for fk in self.inspector.get_foreign_keys(table_name):
                foreign_keys.append({
                    "name": fk['name'],
                    "referred_table": fk['referred_table'],
                    "referred_columns": fk['referred_columns'],
                    "constrained_columns": fk['constrained_columns']
                })

            # Indexes
            indices = []
            for index in self.inspector.get_indexes(table_name):
                indices.append({
                    "name": index['name'],
                    "columns": index['column_names'],
                    "unique": index['unique']
                })

            return {
                "table_name": table_name,
                "columns": columns,
                "primary_key": primary_keys,
                "foreign_keys": foreign_keys,
                "indices": indices
            }
        except SQLAlchemyError as e:
            return f"Failed to get schema for table '{table_name}': {str(e)}"

    def execute_query(self, query, params=None):
        """Execute a custom query"""
        try:
            with self.engine.connect() as connection:
                if params:
                    result = connection.execute(query, params)
                else:
                    result = connection.execute(query)

                # Check if it's a SELECT query
                if result.returns_rows:
                    columns = result.keys()
                    rows = [dict(zip(columns, row)) for row in result]
                    return {"columns": columns, "rows": rows}
                else:
                    return {"affected_rows": result.rowcount}
        except SQLAlchemyError as e:
            return f"Query execution failed: {str(e)}"