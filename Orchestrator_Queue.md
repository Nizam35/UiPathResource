## Specific Data JSON Schema
```json
{
	  "$id": "https://example.com/person.schema.json",
	  "$schema": "https://json-schema.org/draft/2020-12/schema",
	  "title": "Person",
	  "type": "object",
	  "properties": {
	    "firstName": {
	      "type": "string",
	      "description": "The person's first name."
	    },
	    "lastName": {
	      "type": "string",
	      "description": "The person's last name."
	    },
	    "age": {
	      "description": "Age in years which must be equal to or greater than zero.",
	      "type": "integer",
	      "minimum": 0
	    }
	  }
	}
```
## Output Data JSON Schema
```json
{	
		  "$id": "https://example.com/person.schema.json",
		  "$schema": "https://json-schema.org/draft/2020-12/schema",
		  "title": "Output Generated",
		  "type": "object",
		  "properties": {
		    "result": {
		      "type": "object",
		      "description": "The person's first name."
		    },
		    "IsMailSent": {
		      "type": "boolean",
		      "description": "Is the mail has been sent"
		    }
		    
		  }
		}

```
## Analytics Data JSON Schema
```json
{
	  "$id": "https://example.com/person.schema.json",
	  "$schema": "https://json-schema.org/draft/2020-12/schema",
	  "title": "Analytic Data",
	  "type": "object",
	  "properties": {
	    "result": {
	      "type": "object",
	      "description": "The person's first name."
	    },
	    "Test": {
	      "type": "boolean",
	      "description": "Is the mail has been sent"
	    }
	    
	  }
	}

```

