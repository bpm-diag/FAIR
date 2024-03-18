import requests

def get_doi_bibtex(doi):
  base_url = f"https://doi.org/{doi}"
  headers = {
      "Accept": "text/bibliography; style=bibtex"
  }
  response = requests.get(base_url, headers=headers)

  if response.status_code == 200:
    return response.text.encode('ascii', 'ignore').decode('utf-8').strip()

  else:
    print(response.status_code)
    print(response.text)
    return None

l=["10.1515/9783110676693-011",
   "10.1007/978-3-030-91431-8_5",
   "10.1016/j.is.2023.102180",
   "10.1007/978-3-030-21297-1_2"]

for x in l:
    print("#################")
    print(get_doi_bibtex(x))
print("#################")

