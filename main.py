import pandas as pd
import numpy as np

#считываем данные
df_movies = pd.read_excel('movies.xlsx', sheet_name='Movies')
df_imdb = pd.read_excel('movies.xlsx', sheet_name='IMDb data')
df_superheroes = pd.read_excel('movies.xlsx', sheet_name='Marvel DC')

# Функция для очистки названий
def clean_title(title):
    if pd.isna(title):
        return ''
    title = str(title).lower()
    words_to_remove = ['the ', 'a ', 'an ']
    for word in words_to_remove:
        title = title.replace(word, '')
    title = ''.join(c for c in title if c.isalnum() or c.isspace())
    return title.strip()


#ВОПРОС 9

# Создаем очищенные версии названий
df_movies['clean_title'] = df_movies['Title'].apply(clean_title)
df_imdb['clean_title'] = df_imdb['Title'].apply(clean_title)

# Объединяем таблицы по очищенным названиям
merged_df = pd.merge(df_movies, df_imdb, on='clean_title', how='inner', suffixes=('_main', '_imdb'))

print(f"Found matches: {len(merged_df)}")

# Расчет разницы оценок
if not merged_df.empty:
    merged_df['Vote_average'] = pd.to_numeric(merged_df['Vote_average'], errors='coerce')
    merged_df['IMDB Score'] = pd.to_numeric(merged_df['IMDB Score'], errors='coerce')
    
    valid_ratings = merged_df.dropna(subset=['Vote_average', 'IMDB Score'])
    
    if not valid_ratings.empty:
        valid_ratings['rating_diff'] = abs(valid_ratings['Vote_average'] - valid_ratings['IMDB Score'])
        average_diff = valid_ratings['rating_diff'].mean()
        
        print(f"Average rating difference: {average_diff:.4f}")
        
    else:
        print("No films with both ratings found")
else:
    print("No matches found between tables")


#ВОПРОС 10

# Создаем очищенные названия для супергеройских фильмов
df_superheroes['clean_title'] = df_superheroes['Movie'].apply(clean_title)

# Сопоставляем с основной таблицей
merged_super = pd.merge(df_superheroes, df_movies, left_on='clean_title', right_on='clean_title', how='left')

# Группируем по студии и считаем метрики
studio_stats = merged_super.groupby('Company').agg(
    total_movies=('Movie', 'count'),
    average_rating=('Vote_average', 'mean'),
    average_profit=('Profit', 'mean'),
    total_profit=('Profit', 'sum'),
    rating_std=('Vote_average', 'std')
).round(2)

print("\nStudio statistics:")
print(studio_stats)

# Определяем победителя по каждому критерию
print("\ncomparison_result:")
criteria = {
    'total_movies': 'Total_films',
    'average_rating': 'Avg_rating', 
    'total_profit': 'Total_profit',
    'average_profit': 'Avg_profit_per_film'
}

for criterion, description in criteria.items():
    marvel_value = studio_stats.loc['Marvel', criterion]
    dc_value = studio_stats.loc['DC', criterion]
    
    if criterion in ['total_movies', 'average_rating', 'total_profit', 'average_profit']:
        if marvel_value > dc_value:
            print(f"{description}: Marvel ({marvel_value}) > DC ({dc_value})")
        else:
            print(f"{description}: DC ({dc_value}) > Marvel ({marvel_value})")

# Итоговый вердикт
print(f"\nFINAL VERDICT:")
print(f"Analyzed {len(merged_super)} superhero movies")

if studio_stats.loc['Marvel', 'average_rating'] > studio_stats.loc['DC', 'average_rating']:
    print("Marvel shows higher movies quality")
else:
    print("DC shows higher movies quality")

if studio_stats.loc['Marvel', 'total_profit'] > studio_stats.loc['DC', 'total_profit']:
    print("Marvel is more commercially successful")
else:
    print("DC is more commercially successful")
