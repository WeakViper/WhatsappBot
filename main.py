import requests
import json
import openpyxl



url = "https://graph.facebook.com/v16.0/104540322641740/messages"

headers = {
    "Content-Type": "application/json",
    #"Authorization": "Bearer EAAIjg66wjAMBAOvHZAkuET0fND5jSOwNdP9ZBfI0dbxVSKmD9rHY9W0V4Jc6SXYZAVG8XYFE7lhzu47DBtqiZCXfxWZC3OfB05E40JFOCPi1PCnZAGX8zypKBndZBKcD0zxZBqefqZAozd41IRubc5ZCAYLFhSE6I2mhd9FThFCQizCNpNal29MukHMQRZCuoXr6cZB3ZBWhXDcsg5NZAZBHpKNNaZA586QKinryqdUZD"
}

data = {
    "messaging_product": "whatsapp",
    "to": "12365145087",
    "type": "template",
    "template": { "name": "reminder", "language": { "code": "en_US" },
                  "components": [{
                      "type": "body",
                      "parameters": [{
                          "type": "text",
                          "text": "Saira"
                      },
                          {
                              "type": "text",
                              "text": "300"
                          },
                          {
                              "type": "text",
                              "text": "31st of May"
                          }
                      ]
                  }]
                  }
}

response = requests.post(url, headers=headers, data=json.dumps(data))
print(response.json())
