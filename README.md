# WebScraper
### Takes website links as input and searches for specific data in the websites

# Set Up 
- install `pipenv`
    ```sh
    pip install pipenv
    ```
- install the required dependencies from the `pipfile`
    ```sh
    pipenv install
    ```

# Run
1. ### MainScraper
    - Change the directory to `MainScraper`
    - To list all webpages of a site run
        ```sh
        pipenv run python list_urls.py
        ```
    - To find the codes from a site run
        ```sh
        pipenv run python get_codes.py
        ```
    - To find the necessary data run
        ```sh
        pipenv run python find_data.py
        ```

1. ### DetectInputField
    - Change the directory to `DetectInputField`
    - To find data in input fields run
        ```sh
        pipenv run python search.py
        ```

## Tweaks
- Edit the variable `FILE_PATH` in each file to edit the input excel file path
- Name the excel sheet with website links as `Websites`