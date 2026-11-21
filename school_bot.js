'use strict';

// School report bot. Public figures and public data only.
// Local mode: node school_bot.js (uses Happ proxy if available).
// Cloud mode: set BOT_TOKEN and WEBHOOK_URL; app listens on PORT.

const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { Telegraf, Markup } = require('telegraf');
const PptxGenJS = require('pptxgenjs');
const express = require('express');
const HttpsProxyAgentClass = require('https-proxy-agent');

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) throw new Error('BOT_TOKEN is not set');
const USE_PROXY = !!process.env.USE_PROXY;
const PROXY_URL = process.env.PROXY_URL || 'socks://127.0.0.1:10808';
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.WEBHOOK_URL || '';
const BOT_VERSION = '2026-11-21-v4';

const WIKI_API = 'https://ru.wikipedia.org/w/api.php';
const WIKI_FALLBACK = 'https://ru.m.wikipedia.org/w/api.php';
const WIKIDATA_API = 'https://www.wikidata.org/w/api.php';
const SPARQL_API = 'https://query.wikidata.org/sparql';
const MAX_SEARCH_MS = 29000;
const MAX_REQUEST_MS = 9000;

const http = axios.create({
  timeout: MAX_REQUEST_MS,
  headers: {
    'User-Agent': 'SchoolBot/1.0 (school project; t.me/schol_spg_bot)',
    Accept: 'application/json,text/plain,*/*'
  },
  proxy: false
});
if (USE_PROXY) http.defaults.httpsAgent = new HttpsProxyAgentClass(PROXY_URL);

// ============================ DATA ============================

// Canonical public figures with manually checked, school-safe facts.
const LOCAL_PEOPLE = [
  {
    name: 'Уинстон Черчилль',
    aliases: [
      'чертиль', 'чертиля', 'чертил', 'чертильч', 'черчиль', 'черчильч',
      'черчилль', 'черчилля', 'уинстон чертиль', 'уинстона чертиля',
      'уинстон черчилль', 'уинстона черчилля', 'уинстон черчиль',
      'увинстона чертиля', 'увинстон чертиль', 'wineston churchill',
      'winston churchill', 'уинстон churchill'
    ],
    facts: {
      birth: 1874, death: 1965, born: 'Бленхейм, Англия', country: 'Великобритания',
      job: 'политик, премьер-министр, писатель',
      education: 'Королевская военная академия в Сандхёрсте',
      career: 'офицер, журналист, политик; премьер-министр Великобритании в 1940–1945 и 1951–1955 гг.',
      raised: 'в аристократической семье; отец — лорд Рандольф Черчилль, мать — Дженни Джером',
      family: { 'Отец': ['Рандольф Черчилль'], 'Мать': ['Дженни Джером'], 'Жена': ['Клементина Черчилль'] }
    },
    text: 'Уинстон Черчилль родился 30 ноября 1874 года в Бленхеймском дворце. Учился в Королевской военной академии в Сандхёрсте, затем служил офицером и работал военным корреспондентом. После этого он пошёл в политику: неоднократно избирался в парламент, занимал министерские посты, а в 1940 году стал премьер-министром Великобритании. В годы Второй мировой войны его речи и решения пользовались большой поддержкой в стране. После войны он снова вернулся в политику, а также написал много исторических и мемуарных книг. В 1953 году Черчилль получил Нобелевскую премию по литературе.'
  },
  {
    name: 'Альберт Эйнштейн',
    aliases: ['энштейн', 'энштейна', 'ейнштейн', 'айнштейн', 'einstein', 'альберт энштейн', 'альберта энштейна'],
    facts: {
      birth: 1879, death: 1955, born: 'Ульм, Германия', country: 'Германия / США',
      job: 'физик-теоретик',
      education: 'школа в Мюнхене, Цюрихский политехникум',
      career: 'работал в патентном бюро, затем профессором и исследователем; создал теорию относительности',
      raised: 'в семье инженера и предпринимателя; дядя поощрял интерес к науке',
      family: { 'Отец': ['Герман Эйнштейн'], 'Мать': ['Паулина Кох'], 'Жена': ['Милева Марич', 'Эльза Левенталь'] }
    },
    text: 'Альберт Эйнштейн родился 14 марта 1879 года в Ульме. Детство прошло в Мюнхене, где его семья владела фирмой по производству электрооборудования. После переезда в Италию он продолжил обучение в Швейцарии и окончил Цюрихский политехникум. Некоторое время работал в патентном бюро, а затем занялся научной деятельностью. В 1905 году он опубликовал работы по специальной теории относительности, фотоэффекту и броуновскому движению. Позже создал общую теорию относительности. В 1921 году получил Нобелевскую премию по физике. В 1930-х годах переехал в США и стал одним из самых известных учёных XX века.'
  },
  {
    name: 'Александр Пушкин',
    aliases: ['пушкин', 'пушкина', 'александр пушкин', 'александра пушкина', 'саша пушкин'],
    facts: {
      birth: 1799, death: 1837, born: 'Москва', country: 'Россия',
      job: 'поэт, прозаик, драматург',
      education: 'Лицей в Царском Селе',
      career: ' начинал как поэт, служил в Министерстве иностранных дел, затем полностью посвятил себя литературе',
      raised: 'в дворянской семье; воспитанием и образованием занимались гувернёры и наставники лицея',
      family: { 'Отец': ['Сергей Пушкин'], 'Мать': ['Надежда Ганнибал'], 'Жена': ['Наталья Гончарова'] }
    },
    text: 'Александр Сергеевич Пушкин родился 26 мая 1799 года в Москве. Отец происходил из старинного дворянского рода, мать была внучкой африканца Ибрагима Ганнибала. В 1811 году Пушкин поступил в только что открытый Лицей в Царском Селе, где получил отличное образование. После окончания лицея он служил в Министерстве иностранных дел, но главным делом жизни стала литература. Пушкин создал роман «Евгений Онегин», трагедию «Борис Годунов», «Капитанскую дочку», много стихотворений и сказок. Он считается основоположником современной русской литературной формы.'
  },
  {
    name: 'Юрий Гагарин',
    aliases: ['гагарин', 'гагарина', 'юра гагарин', 'первый космонавт'],
    facts: {
      birth: 1934, death: 1968, born: 'Клушино, Смоленская область', country: 'СССР',
      job: 'космонавт, лётчик',
      education: 'Саратовский индустриальный техникум, Оренбургское военное авиационное училище',
      career: ' стал лётчиком, затем прошёл подготовку в отряде космонавтов; 12 апреля 1961 года совершил первый в истории полёт в космос',
      raised: 'в семье колхозников; после войны семья переехала в Гжатск',
      family: { 'Отец': ['Алексей Гагарин'], 'Мать': ['Анна Гагарина'], 'Жена': ['Валентина Гагарина'] }
    },
    text: 'Юрий Алексеевич Гагарин родился 9 марта 1934 года в деревне Клушино. Учился в Саратовском индустриальном техникуме, затем в Оренбургском авиационном училище. Пошёл по стопам старшего брата и стал лётчиком. В 1960 году был зачислен в первый отряд космонавтов. 12 апреля 1961 года на корабле «Восток-1» он впервые в истории человека совершил орбитальный полёт вокруг Земли. После полёта стал всемирно известным, занимался общественной деятельностью и подготовкой новых космонавтов.'
  },
  {
    name: 'Лев Толстой',
    aliases: ['толстой', 'толстого', 'лев толстой', 'льва толстого', 'граф толстой'],
    facts: {
      birth: 1828, death: 1910, born: 'Ясная Поляна, Тульская губерния', country: 'Россия',
      job: 'писатель, мыслитель',
      education: 'домашнее образование, затем Казанский университет',
      career: 'участвовал в Крымской войне, затем занимался литературой и педагогикой',
      raised: 'в дворянской семье; получил домашнее воспитание',
      family: { 'Отец': ['Николай Толстой'], 'Мать': ['Мария Волконская'], 'Жена': ['София Толстая'] }
    },
    text: 'Лев Николаевич Толстой родился в 1828 году в Ясной Поляне. Получил домашнее образование, затем учился в Казанском университете. После военной службы занялся литературой. Главные произведения — романы «Война и мир», «Анна Каренина», «Воскресение». В последние годы жизни Толстой отказался от дворянского образа жизни, проповедовал ненасилие и простоту. Его творчество оказало огромное влияние на русскую и мировую литературу.'
  },
  {
    name: 'Наполеон Бонапарт',
    aliases: ['наполеон', 'наполеона', 'бонапарт', 'бонапарта', 'наполеон бонапарт', 'наполеона бонапарта'],
    facts: {
      birth: 1769, death: 1821, born: 'Аяччо, Корсика', country: 'Франция',
      job: 'полководец, император',
      education: 'военная школа в Бриенне, Парижская военная школа',
      career: 'военный офицер, затем консул и император Франции',
      raised: 'в небогатой дворянской семье на Корсике; воспитывался в военных школах',
      family: { 'Отец': ['Карло Бонапарт'], 'Мать': ['Летиция Рамолино'], 'Жена': ['Жозефина', 'Мария Луиза'] }
    },
    text: 'Наполеон Бонапарт родился в 1769 году в Аяччо на Корсике. Учился в военных школах во Франции и рано начал военную службу. Во время Французской революции быстро выдвинулся, затем стал консулом, а в 1804 году — императором. Он провёл много военных кампаний и создал крупную империю, но после поражения в 1814 году отрёкся от престола. В 1815 году попытался вернуться к власти, но потерпел поражение при Ватерлоо и был сослан на остров Святой Елены.'
  },
  {
    name: 'Мария Кюри',
    aliases: ['мария кюри', 'марию кюри', 'кюри', 'мари курi', 'curie', 'мария склодовская-кюри'],
    facts: {
      birth: 1867, death: 1934, born: 'Варшава', country: 'Польша / Франция',
      job: 'физик и химик',
      education: 'Сорбонна в Париже',
      career: 'исследователь радиоактивности, профессор Сорбонны',
      raised: 'в семье учителей; с детства проявляла интерес к естественным наукам',
      family: { 'Отец': ['Владислав Склодовский'], 'Мать': ['Бронислава Склодовская'], 'Муж': ['Пьер Кюри'] }
    },
    text: 'Мария Склодовская-Кюри родилась в 1867 году в Варшаве. Училась в подпольных польских школах, затем в Сорбонне в Париже. вместе с мужем Пьером Кюри исследовала радиоактивные элементы и открыла полоний и радий. Дважды получала Нобелевскую премию: по физике в 1903 году и по химии в 1911 году. Она стала первой женщиной — профессором Сорбонны.'
  },
  {
    name: 'Михаил Ломоносов',
    aliases: ['ломоносов', 'ломоносова', 'михаил ломоносов', 'михаила ломоносова'],
    facts: {
      birth: 1711, death: 1765, born: 'Денисовка, Архангельская губерния', country: 'Россия',
      job: 'учёный, поэт',
      education: 'Славяно-греко-латинская академия в Москве, университет в Марбурге',
      career: 'профессор химии, основатель Московского университета',
      raised: 'в семье поморского крестьянина; с юности стремился к учёбе',
      family: { 'Отец': ['Василий Ломоносов'], 'Мать': ['Елена Ломоносова'], 'Жена': ['Елизавета Дильбах'] }
    },
    text: 'Михаил Васильевич Ломоносов родился в 1711 году в поморской деревне. В 19 лет пешком отправился в Москву и поступил в Славяно-греко-латинскую академию. Затем учился в Германии. Вернувшись в Россию, стал профессором, занимался химией, физикой, геологией и поэзией. Вместе с Шуваловым основал Московский университет. Его работы заложили основы русской научной традиции.'
  },
  {
    name: 'Иван Грозный',
    aliases: ['иван грозный', 'ивана грозного', 'грозный', 'иван iv', 'иван васильевич'],
    facts: {
      birth: 1530, death: 1584, born: 'Москва', country: 'Русское царство',
      job: 'царь',
      education: 'домашнее образование при дворе',
      career: 'великий князь московский, затем первый русский царь',
      raised: 'при дворе в окружении бояр; детство прошло в условиях придворных конфликтов',
      family: { 'Отец': ['Василий III'], 'Мать': ['Елена Глинская'], 'Жены': ['Анастасия Романова', 'Мария Темрюковна и др.'] }
    },
    text: 'Иван IV Васильевич родился в 1530 году. В 1547 году был коронован как первый русский царь. При нём были проведены реформы, созданы приказы, начата Ливонская война. В памяти народа остался как жестокий правитель эпохи опричнины, но также как создатель централизованного Русского государства.'
  },
  {
    name: 'Пётр I',
    aliases: ['пётр первый', 'петр первый', 'пётр 1', 'петр 1', 'пётр великий', 'петр великий', 'пётр', 'петр алексеевич'],
    facts: {
      birth: 1672, death: 1725, born: 'Москва', country: 'Россия',
      job: 'царь, реформатор',
      education: 'домашнее образование, затем обучение в Голландии и Англии',
      career: 'царь, создал регулярную армию и флот, основал Санкт-Петербург',
      raised: 'в царской семье; с детства интересовался военным делом и мореплаванием',
      family: { 'Отец': ['Алексей Михайлович'], 'Мать': ['Наталья Нарышкина'], 'Жена': ['Екатерина I'] }
    },
    text: 'Пётр I родился в 1672 году. С юности увлекался военным делом и кораблями. В 1690-х годах его называли «царём-плотником». Он провёл реформы армии, управления и образования, основал Санкт-Петербург и Академию наук. После победы в Северной войне Россия вышла к Балтийскому морю. Пётр I считается одним из главных реформаторов в русской истории.'
  },
  {
    name: 'Карл Маркс',
    aliases: ['маркс', 'маркса', 'карл маркс', 'карла маркса'],
    facts: {
      birth: 1818, death: 1883, born: 'Трир, Германия', country: 'Германия',
      job: 'философ, экономист',
      education: 'университеты Бонна и Берлина',
      career: 'журналист и теоретик; создал учение о капитализме вместе с Энгельсом',
      raised: 'в семье адвоката; получил классическое образование',
      family: { 'Отец': ['Генрих Маркс'], 'Мать': ['Генриетта Прево'], 'Жена': ['Женни фон Вестфален'] }
    },
    text: 'Карл Маркс родился в 1818 году в Трире. Учился в университетах Бонна и Берлина. Занимался журналистикой, но из-за своих взглядов не смог работать в Германии и переехал в Париж, затем в Лондон. Вместе с Фридрихом Энгельсом написал «Манифест Коммунистической партии» и «Капитал». Его идеи оказали огромное влияние на историю XX века.'
  },
  {
    name: 'Фридрих Энгельс',
    aliases: ['энгельс', 'энгельса', 'фридрих энгельс', 'фридриха энгельса'],
    facts: {
      birth: 1820, death: 1895, born: 'Бармен, Германия', country: 'Германия',
      job: 'философ, экономист',
      education: 'гимназия, коммерческое училище',
      career: 'бизнесмен, журналист и теоретик; соавтор Маркса',
      raised: 'в семье фабриканта',
      family: { 'Отец': ['Фридрих Энгельс-старший'], 'Мать': ['Элизабет ван Хорн'] }
    },
    text: 'Фридрих Энгельс родился в 1820 году. Работал в торговой фирме в Манчестере и изучал положение рабочих. Вместе с Карлом Марксом написал «Манифест Коммунистической партии» и другие работы. После смерти Маркса подготовил к печати второй и третий тома «Капитала».'
  },
  {
    name: 'Уильям Шекспир',
    aliases: ['шекспир', 'шекспира', 'уильям шекспир', 'вильям шекспир', 'shakespeare'],
    facts: {
      birth: 1564, death: 1616, born: 'Стратфорд-на-Эйвоне', country: 'Англия',
      job: 'драматург, поэт',
      education: 'школа в Стратфорде',
      career: 'актёр, драматург, владелец театра «Глобус»',
      raised: 'в семье перчаточника и мелкого чиновника',
      family: { 'Отец': ['Джон Шекспир'], 'Мать': ['Мэри Арден'], 'Жена': ['Анна Хэтэуэй'] }
    },
    text: 'Уильям Шекспир родился в 1564 году в Стратфорде. Учился в местной школе, затем стал актёром и драматургом в Лондоне. Написал 37 пьес, среди них «Гамлет», «Макбет», «Король Лир», «Ромео и Джульетта». Его произведения переведены на многие языки и до сих пор ставятся в театрах.'
  },
  {
    name: 'Чарльз Дарвин',
    aliases: ['дарвин', 'дарвина', 'чарльз дарвин', 'чарлз дарвин', 'darwin'],
    facts: {
      birth: 1809, death: 1882, born: 'Шрусбери, Англия', country: 'Великобритания',
      job: 'естествоиспытатель',
      education: 'школа в Шрусбери, Эдинбургский и Кембриджский университеты',
      career: 'путешественник и исследователь; создал теорию эволюции',
      raised: 'в семье врача; с детства любил собирать коллекции',
      family: { 'Отец': ['Роберт Дарвин'], 'Мать': ['Сюзанна Уэджвуд'], 'Жена': ['Эмма Уэджвуд'] }
    },
    text: 'Чарльз Дарвин родился в 1809 году. Учился в школе, затем в университетах. В 1831 году отправился в плавание на корабле «Бигль», что дало ему материал для наблюдений. В 1859 году опубликовал книгу «Происхождение видов», где объяснил теорию эволюции естественным отбором. Его идеи изменили биологию.'
  },
  {
    name: 'Александр Невский',
    aliases: ['александр невский', 'александра невского', 'невский', 'невского'],
    facts: {
      birth: 1220, death: 1263, born: 'Переславль-Залесский', country: 'Древняя Русь',
      job: 'князь, полководец',
      education: 'домашнее образование',
      career: 'новгородский князь; победил шведов на Неве и крестоносцев на Чудском озере',
      raised: 'в семье князя Ярослава Всеволодовича',
      family: { 'Отец': ['Ярослав Всеволодович'], 'Мать': ['Феодосия'], 'Жена': ['Васса'] }
    },
    text: 'Александр Невский родился в 1220 году. Был новгородским князем. В 1240 году разбил шведское войско на реке Неве, а в 1242 году — рыцарей на Чудском озере. При нём Новгород сохранил независимость, а сам он стал символом защиты русских земель. Признан святым Русской православной церковью.'
  },
  {
    name: 'Дмитрий Менделеев',
    aliases: ['менделеев', 'менделеева', 'дмитрий менделеев', 'дмитрия менделеева'],
    facts: {
      birth: 1834, death: 1907, born: 'Тобольск', country: 'Россия',
      job: 'химик',
      education: 'Тобольская гимназия, Петербургский университет',
      career: 'профессор химии; открыл периодический закон и создал таблицу элементов',
      raised: 'в большой семье директора гимназии',
      family: { 'Отец': ['Иван Менделеев'], 'Мать': ['Мария Корнилова'], 'Жена': ['Анна Попова'] }
    },
    text: 'Дмитрий Иванович Менделеев родился в 1834 году в Тобольске. Учился в гимназии и Петербургском университете. В 1869 году открыл периодический закон химических элементов и составил периодическую таблицу. Также занимался вопросами промышленности и метрологии.'
  },
  {
    name: 'Исаак Ньютон',
    aliases: ['ньютон', 'ньютона', 'исаак ньютон', 'исаака ньютона', 'newton'],
    facts: {
      birth: 1643, death: 1727, born: 'Вулсторп, Англия', country: 'Великобритания',
      job: 'физик, математик',
      education: 'Тринити-колледж Кембриджского университета',
      career: 'профессор математики; создал классическую механику и теорию тяготения',
      raised: 'в семье фермера; воспитывался бабушкой после смерти матери',
      family: { 'Отец': ['Исаак Ньютон-старший'], 'Мать': ['Ханна Эйскью'] }
    },
    text: 'Исаак Ньютон родился в 1643 году. Учился в Кембридже. Сформулировал законы движения и закон всемирного тяготения, создал основы дифференциального исчисления. Его книга «Математические начала натуральной философии» стала одним из главных научных трудов в истории.'
  },
  {
    name: 'София Ковалевская',
    aliases: ['ковалевская', 'ковалевской', 'софия ковалевская', 'софья ковалевская', 'софии ковалевской'],
    facts: {
      birth: 1850, death: 1891, born: 'Москва', country: 'Россия / Швеция',
      job: 'математик',
      education: 'Гейдельбергский и Берлинский университеты',
      career: 'первая женщина — профессор математики в Северной Европе',
      raised: 'в дворянской семье; интерес к математике проявился в детстве',
      family: { 'Отец': ['Василий Корвин-Круковский'], 'Мать': ['Елизавета Шуберт'], 'Муж': ['Владимир Ковалевский'] }
    },
    text: 'София Ковалевская родилась в 1850 году в Москве. Из-за запрета женщинам учиться в российских университетах она уехала в Германию и слушала лекции в Гейдельберге и Берлине. Стала первой женщиной, получившей докторскую степень по математике, и первой женщиной-профессором в Стокгольмском университете.'
  },
  {
    name: 'Владимир Ленин',
    aliases: ['ленин', 'ленина', 'владимир ленин', 'владимира ленина', 'ульянов', 'ульянова'],
    facts: {
      birth: 1870, death: 1924, born: 'Симбирск', country: 'Россия / СССР',
      job: 'политик, революционер',
      education: 'Симбирская гимназия, Казанский университет',
      career: 'революционер, основатель большевистской партии, первый глава советского государства',
      raised: 'в семье учителя; старший брат был казнён, что повлияло на его взгляды',
      family: { 'Отец': ['Илья Ульянов'], 'Мать': ['Мария Бланк'], 'Жена': ['Надежда Крупская'] }
    },
    text: 'Владимир Ильич Ленин родился в 1870 году в Симбирске. Учился в гимназии и Казанском университете. После казни старшего брата увлёкся революционными идеями. В 1917 году возглавил большевиков и Октябрьскую революцию. С 1917 года возглавлял советское правительство.'
  },
  {
    name: 'Александр Суворов',
    aliases: ['суворов', 'суворова', 'александр суворов', 'александра суворова'],
    facts: {
      birth: 1730, death: 1800, born: 'Москва', country: 'Россия',
      job: 'полководец',
      education: 'домашнее образование',
      career: 'военный командир; провёл Итальянский и Швейцарский походы',
      raised: 'в семье военного; с детства готовился к службе',
      family: { 'Отец': ['Василий Суворов'], 'Мать': ['Мария Манукова'] }
    },
    text: 'Александр Васильевич Суворов родился в 1730 году. С детства готовился к военной службе. Отличился в Русско-турецких войнах, при взятии Измаила, в Итальянском и Швейцарском походах. Известен как выдающийся полководец, никогда не проигрывавший сражения.'
  }
];

const ALIASES = {};
for (const p of LOCAL_PEOPLE) {
  ALIASES[norm(p.name)] = p.name;
  for (const a of p.aliases) ALIASES[norm(a)] = p.name;
}

const STOP_WORDS = new Set([
  'привет','здравствуй','здравствуйте','хай','добрый','день','утро','вечер',
  'мне','надо','нужно','нжен','нужна','хочу','хочет','пожалуйста','пжлст','пж',
  'напиши','написать','сделай','сделать','сгенерируй','создай','давай','дай','пиши',
  'про','о','об','обо','вая','во','информация','инфу','инфо','сведения','расскажи',
  'рассказать','даные','данные',' доклад','реферат','сочинение','текст','письменная',
  'работа',' докладчик','презентация','презентации','презент','през','слайды','слайд',
  'сделай','сделать','на tema','тема','тему','это','этот','эта','эти','тот','та','того',
  'какой','какая','какие','кто','что','как','где','когда','почему','зачем','для',
  'школьный','школа',' clas','класс','класса','ученик','ученица','учитель',
  'хороший','красивый','красиво','подробно','подробная','подробный','коротко',
  'кратко','просто','попроще','серьёзно','серьезно','быстро','ещё','еще','и','а','но',
  'или','да','нет','не','без','без','же','бы','в','во','на','за','из','к','ко','у','от',
  'до','при','по','с','со','вот','уже','тоже','так','итак','например','например',
  'сам','сама','сами','себя','его','её','ее','их','його','им','ому','ему','ей','ихний',
  'нд','нде','н надо','ненадо','не надо','не надо','ненадо','ненад','ненада',
  'семью','семя','семьи','семей','семью','жену','мужа','детень','дети','детей',
  'войну','война','про войну','военный','военные','боевые','боевые действия',
  'учился','училась','где учился','где он учился','где она училась','образование',
  'пошёл','пошла','на кого пошёл','путь','карьера','деятельность','род деятельности',
  'воспитывал','воспитывала','кто воспитывал','родители','мама','папа','бабушка',
  'дедушка','род','родом','родился','родилась','биография','жизнь',' открытия',
  'вывод','заключение','введение','основная часть','глава','пункт','список',
  'можешь','сможешь',' есть','имеется','найди','найти','поищи','проверь','сверь',
  'источники','источник','Википедия','Вики','Так','У','КБЖД','готовый','готовое',
  'вариант','пример','шаблон','structure','структура','по пунктам','по пунктам'
]);

const RU_MAP = {
  а:['a','а'], б:['b','б'], в:['v','в'], г:['g','г'], д:['d','д'], е:['e','е'], ё:['e','ё'],
  ж:['zh','ж'], з:['z','з'], и:['i','и'], й:['y','й'], к:['k','к'], л:['l','л'], м:['m','м'],
  н:['n','н'], о:['o','о'], п:['p','п'], р:['r','р'], с:['s','с'], т:['t','т'], у:['u','у'],
  ф:['f','ф'], х:['h','х'], ц:['c','ц'], ч:['ch','ч'], ш:['sh','ш'], щ:['sch','щ'],
  ъ:['','ъ'], ы:['y','ы'], ь:['','ь'], э:['e','э'], ю:['u','ю'], я:['ya','я']
};
const TRANSLIT = {
  a:'а', b:'б', v:'в', g:'г', d:'д', e:'е', zh:'ж', z:'з', i:'и', y:'й', k:'к', l:'л',
  m:'м', n:'н', o:'о', p:'п', r:'р', s:'с', t:'т', u:'у', f:'ф', h:'х', c:'ц', ch:'ч',
  sh:'ш', sch:'щ', w:'в', x:'кс'
};

// ============================ UTIL ============================

function dot(s) {
  return String(s || '').replace(/\.+$/, '') + '.';
}

function cap(s) {
  const v = String(s || '').trim();
  return v ? v[0].toUpperCase() + v.slice(1) : v;
}

function norm(s) {
  return String(s || '').toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripWordEndings(token) {
  // Remove common Russian inflections for name search, keep enough for fuzzy match.
  return token
    .replace(/(ого|его|ому|ему|ыми|ими|ами|ями|ах|ях|ов|ев|ой|ей|ом|ем|у|ю|а|я|ы|и|е|ь)$/g, '');
}

function levenshtein(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const prev = new Array(b.length + 1);
  const cur = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    cur[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        cur[j] = Math.min(cur[j], prev[j - 2] + 1);
      }
    }
    for (let j = 0; j <= b.length; j++) prev[j] = cur[j];
  }
  return prev[b.length];
}

function t9Variants(word) {
  // If Latin typed in Russian layout or vice versa, transliterate both ways.
  const w = norm(word).replace(/\s/g, '');
  const cyr = /[а-я]/.test(w);
  let out = '';
  for (const ch of w) {
    if (/[a-z]/.test(ch)) {
      out += TRANSLIT[ch] || ch;
    } else if (RU_MAP[ch]) {
      out += RU_MAP[ch][0];
    } else out += ch;
  }
  return [w, out].filter(Boolean);
}

function similarity(a, b) {
  if (!a || !b) return 0;
  const max = Math.max(a.length, b.length);
  return Math.max(0, 1 - levenshtein(a, b) / max);
}

function bestNameScore(candidate, targetTokens) {
  const candTokens = norm(candidate).split(' ').filter(Boolean);
  let total = 0;
  let used = 0;
  for (const ct of candTokens) {
    const stem = stripWordEndings(ct);
    let best = 0;
    for (const tt of targetTokens) {
      for (const v of t9Variants(tt)) {
        const vStem = stripWordEndings(v);
        const s = Math.max(similarity(ct, v), similarity(stem, vStem), similarity(stem, v));
        if (s > best) best = s;
      }
    }
    total += best;
    if (best > 0.5) used++;
  }
  const coverage = used / Math.max(1, candTokens.length);
  const extraPenalty = Math.max(0, targetTokens.length - candTokens.length) * 0.08;
  return total / Math.max(1, candTokens.length) * 0.72 + coverage * 0.28 - extraPenalty;
}

// Parse all user constraints from the entire message.
function parseRequest(raw, mode) {
  const original = String(raw || '').trim();
  const lower = norm(original);
  const constraints = [];
  const add = (flag, phrase) => { if (flag && !constraints.includes(phrase)) constraints.push(phrase); };

  const noFamily = /без\s+сем(?:ью|ьи|и|я)|без\s+[а-яё]+\s+и\s+сем(?:ью|ьи|и|я)|сем(?:ью|ьи|и|я)\s*н[еи]?\s*над|н[еи]?\s*над[оа]?\s+сем(?:ью|ьи|и|я)|без\s+жен|без\s+муж|без\s+дет|сем(?:ью|ьи|и|я)\s+не\s+пиши|не\s+пиши\s+сем(?:ью|ьи|и|я)|сем(?:ью|ьи|и|я)\s+не\s+нуж|без\s+родств/.test(lower);
  const noWar = /без\s+во[йи]н|во[йи]н[уы]\s*н[еи]?\s*над|н[еи]?\s*над[оа]?\s*(?:про\s+)?во[йи]н|как\s+про\s+во[йи]н|не\s+пиши\s+во[йи]н|без\s+боев|без\s+воен|военн[а-я]*\s*н[еи]?\s*над/.test(lower);
  const noPolitics = /без\s+полит|политику\s+не\s+над|не\s+над[оа]?\s+полит/.test(lower);
  const onlySchool = /где\s+он\s+учил|где\s+она\s+учил|где\s+учил|где\s+училась|школ|учеб|образован|на\s+кого\s+пош|кем\s+стал|карьер|путь|деятельност/.test(lower);
  const raisedExplicit = /кто\s+воспит|воспит|родител|бабушк|дедушк|детств/.test(lower);
  const raisedSoft = /мам|пап|сем(?:ью|ьи|и|я)/.test(lower) && !noFamily;
  const onlyRaised = raisedExplicit || raisedSoft;
  const grade = (original.match(/(?:^|\s)(\d{1,2})\s*(?:класс|кл|класса)(?:\s|$)/i) || [])[1] || '';

  add(noFamily, 'Без раздела «Семья»');
  add(noWar, 'Без раздела войны и боевых действий');
  add(noPolitics, 'Без политики');
  add(onlySchool, 'Особое внимание: где учился и на кого пошёл');
  add(onlyRaised, 'Особое внимание: кто воспитывал и детство');

  return { original, lower, constraints, grade, noFamily, noWar, noPolitics, onlySchool, onlyRaised, mode };
}

function cleanNameQuery(raw, parsed) {
  let s = norm(raw);

  // Remove buttons/commands.
  s = s.replace(/^(\/start|📝\s*готовый\s*текст|📊\s*презентация)\s*/i, '');

  // Remove explicit constraint phrases before name extraction.
  s = s
    .replace(/без\s+сем[ьиюя][^\n]*/g, ' ')
    .replace(/семью\s+не\s+над[оа]?[^\n]*/g, ' ')
    .replace(/не\s+над[оа]?\s+сем[^\n]*/g, ' ')
    .replace(/без\s+во[йи]н[^\n]*/g, ' ')
    .replace(/войну\s+не\s+над[оа]?[^\n]*/g, ' ')
    .replace(/не\s+над[оа]?\s+во[йи]н[^\n]*/g, ' ')
    .replace(/как\s+про\s+во[йи]н[^\n]*/g, ' ')
    .replace(/без\s+полит[^\n]*/g, ' ')
    .replace(/где\s+он\s+учил[^\n]*/g, ' ')
    .replace(/где\s+она\s+учил[^\n]*/g, ' ')
    .replace(/где\s+учил[^\n]*/g, ' ')
    .replace(/на\s+кого\s+пош[ёе]л[^\n]*/g, ' ')
    .replace(/кто\s+воспитывал[^\n]*/g, ' ')
    .replace(/кто\s+его\s+воспитывал[^\n]*/g, ' ')
    .replace(/кто\s+её\s+воспитывал[^\n]*/g, ' ')
    .replace(/и\s+так\s+далее/g, ' ')
    .replace(/и\s+т\s*д/g, ' ')
    .replace(/так\s+далее/g, ' ')
    .replace(/\b\d{1,2}\s*(?:класс|кл|класса)\b/g, ' ')
    .replace(/\b(?:для|в)\s*\d{1,2}\s*(?:класс|кл|класса)\b/g, ' ');

  // Remove request/action words, keep only likely name words.
  const tokens = s.split(/\s+/).filter(Boolean);
  const kept = [];
  for (const t of tokens) {
    if (STOP_WORDS.has(t)) continue;
    if (/^\d+$/.test(t)) continue;
    const stem = stripWordEndings(t);
    if (STOP_WORDS.has(stem)) continue;
    if (t.length <= 2 && !/[а-яa-z]/.test(t)) continue;
    kept.push(t);
  }

  return kept.join(' ').trim();
}

function findLocal(query) {
  const q = norm(query);
  if (!q) return null;

  // Exact alias wins.
  if (ALIASES[q]) return LOCAL_PEOPLE.find(p => p.name === ALIASES[q]);

  // Also allow alias with ending (например, «пушкина»).
  for (const [alias, canonical] of Object.entries(ALIASES)) {
    const a = stripWordEndings(alias);
    const qq = stripWordEndings(q);
    if (a.length >= 4 && qq.length >= 4 && (a === qq || similarity(a, qq) > 0.9)) {
      return LOCAL_PEOPLE.find(p => p.name === canonical);
    }
  }

  const tokens = q.split(' ').filter(Boolean);
  let best = null;
  let bestScore = 0;
  for (const p of LOCAL_PEOPLE) {
    const candidateVariants = [p.name, ...p.aliases];
    let score = 0;
    for (const v of candidateVariants) {
      const vt = norm(v).split(' ').filter(Boolean);
      score = Math.max(score, bestNameScore(vt, tokens));
    }
    if (score > bestScore) { bestScore = score; best = p; }
  }

  // Require confident match; otherwise let online sources try.
  return bestScore >= 0.67 ? best : null;
}

// ============================ SOURCES ============================

const apiCache = new Map();
let apiQueue = Promise.resolve();
let lastCall = 0;

function apiGet(url, params, opts = {}) {
  const key = url + JSON.stringify(params || {});
  if (apiCache.has(key)) return apiCache.get(key);
  const run = apiQueue.then(async () => {
    const wait = Math.max(0, 380 - (Date.now() - lastCall));
    if (wait) await new Promise(r => setTimeout(r, wait));
    lastCall = Date.now();
    let lastErr = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const res = await http.get(url, { params: params || {}, timeout: opts.timeout || MAX_REQUEST_MS });
        apiCache.set(key, res.data);
        return res.data;
      } catch (e) {
        lastErr = e;
        const status = e && e.response && e.response.status;
        if ((status !== 429 && status < 500) || attempt === 1) break;
        const retryAfter = Number(e.response.headers['retry-after']) || 0;
        await new Promise(r => setTimeout(r, Math.min(2500, retryAfter ? retryAfter * 1000 : 700 * (attempt + 1))));
      }
    }
    throw lastErr;
  });
  apiQueue = run.catch(() => {});
  return run;
}

function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(`${label} timeout`)), ms))
  ]);
}

async function wikiSearch(base, name) {
  const data = await apiGet(base, {
    action: 'query', list: 'search', srsearch: name, srlimit: 6,
    format: 'json', origin: '*'
  });
  return (data.query && data.query.search || []).map(x => x.title);
}

async function wikiExtract(base, title) {
  const data = await apiGet(base, {
    action: 'query', prop: 'extracts|pageprops', exintro: 1, explaintext: 1,
    titles: title, redirects: 1, format: 'json', origin: '*'
  });
  const pages = data.query && data.query.pages || {};
  const page = Object.values(pages)[0] || {};
  return {
    title: page.title || title,
    text: String(page.extract || '').replace(/\s+\n/g, '\n').trim(),
    wikidata: page.pageprops && page.pageprops.wikibase_item || ''
  };
}

async function wikidataLabel(qid) {
  if (!qid) return '';
  try {
    const data = await apiGet(WIKIDATA_API, {
      action: 'wbgetentities', ids: qid, props: 'labels', languages: 'ru|en', format: 'json', origin: '*'
    });
    const ent = data.entities && data.entities[qid];
    return (ent && ent.labels && (ent.labels.ru || ent.labels.en) || {}).value || '';
  } catch (_) { return ''; }
}

async function ddgSearch(name) {
  try {
    const data = await apiGet('https://api.duckduckgo.com/', {
      q: name, format: 'json', no_html: 1, skip_disambig: 1
    });
    const text = [data.AbstractText, data.Answer].filter(Boolean).join('\n');
    return { title: data.Heading || name, text: text.trim() };
  } catch (_) { return null; }
}

function looksPerson(text, title) {
  const t = `${title || ''} ${String(text || '').slice(0, 800)}`;
  const bad = /КБЖД|железная дорога|станция|поезд|маршрут|остановка|автобус| Уикипедия /i;
  if (bad.test(t)) return false;
  const born = /родился|родилась|birth|born|\b18\d\d\b|\b19\d\d\b|\b20\d\d\b/i.test(t);
  const human = /человек|политик|писатель|поэт|учён|учен|физик|химик|биолог|президент|премьер|корол|царь|князь|генерал|художник|композитор|актёр|актер|певец|космонавт|философ|математик|инженер|изобретатель|деятел|персон/i.test(t);
  return born || human;
}

function deriveFacts(title, text) {
  const facts = { birth: '', death: '', born: '', country: '', job: '', education: '', career: '', raised: '', family: {} };
  const b = text.match(/(?:родился|родилась)[^.]{0,120}?(\d{4})/i);
  if (b) facts.birth = b[1];
  const bornPlace = text.match(/(?:родился|родилась)[^.]{0,120}?\bв\s+([^,.]{2,60})/i);
  if (bornPlace) facts.born = bornPlace[1].trim();
  const country = text.match(/(гражданство|государство|страна)[^:]{0,40}:\s*([^.\n]{2,50})/i);
  if (country) facts.country = country[2].trim();
  const job = text.match(/[^.]{0,80}(?:является|был|была| работает|работал|работала)\s+([^,.]{3,60})/i);
  if (job) facts.job = job[1].trim();
  const edu = text.match(/(?:учился|училась|окончил|окончила|образование)[^.]{0,160}/i);
  if (edu) facts.education = edu[0].trim();
  const career = text.match(/(?:карьера|деятельность|работал|работала|занимался|занималась)[^.]{0,180}/i);
  if (career) facts.career = career[0].trim();
  const father = text.match(/отец[^\n.]{0,80}/i);
  const mother = text.match(/мать[^\n.]{0,80}/i);
  if (father) facts.family['Отец'] = [father[0].replace(/^отец\s*/i, '').trim()];
  if (mother) facts.family['Мать'] = [mother[0].replace(/^мать\s*/i, '').trim()];
  return facts;
}

async function collectPerson(rawQuery, parsed) {
  const query = cleanNameQuery(rawQuery, parsed);
  if (!query) return null;

  const local = findLocal(query);
  if (local) return { ...local, title: local.name, source: 'проверенная локальная база' };

  const deadline = Date.now() + MAX_SEARCH_MS;
  const msLeft = () => Math.max(1000, deadline - Date.now());

  const bases = [WIKI_API, WIKI_FALLBACK];
  let wikiHit = null;

  for (const base of bases) {
    if (Date.now() > deadline) break;
    try {
      const titles = await withTimeout(wikiSearch(base, query), Math.min(10000, msLeft()), 'wiki search');
      for (const title of titles.slice(0, 4)) {
        if (Date.now() > deadline) break;
        if (/список|категория|шаблон|портал/i.test(title)) continue;
        const page = await withTimeout(wikiExtract(base, title), Math.min(10000, msLeft()), 'wiki extract');
        if (page.text && looksPerson(page.text, page.title)) {
          wikiHit = { ...page, source: 'Википедия' };
          break;
        }
      }
    } catch (_) {}
    if (wikiHit) break;
  }

  let wdHit = null;
  if (wikiHit && wikiHit.wikidata && Date.now() < deadline) {
    const label = await withTimeout(wikidataLabel(wikiHit.wikidata), Math.min(6000, msLeft()), 'wikidata').catch(() => '');
    if (label && looksPerson(label, label)) wdHit = label;
  }

  let ddgHit = null;
  if ((!wikiHit || !looksPerson(wikiHit.text, wikiHit.title)) && Date.now() < deadline) {
    ddgHit = await withTimeout(ddgSearch(query), Math.min(8000, msLeft()), 'ddg').catch(() => null);
  }

  const candidates = [];
  if (wikiHit) candidates.push({ title: wikiHit.title, text: wikiHit.text, source: 'Википедия', score: 3 });
  if (wdHit) candidates.push({ title: wdHit, text: '', source: 'Wikidata', score: 1 });
  if (ddgHit && ddgHit.text && looksPerson(ddgHit.text, ddgHit.title)) {
    candidates.push({ title: ddgHit.title || query, text: ddgHit.text, source: 'DuckDuckGo', score: 2 });
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) => b.score - a.score);
  const best = candidates[0];
  return {
    name: best.title,
    title: best.title,
    text: best.text || `${best.title} — общественная личность. Подробная биография доступна в открытых источниках.`,
    facts: deriveFacts(best.title, best.text),
    source: candidates.map(c => c.source).join(' + ')
  };
}

// ============================ REPORT ============================

function makeReport(d, note, parsed) {
  const f = d.facts || {};
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. ${f.birth ? 'Родился(ась) в ' + f.birth + ' году' : 'Информация о годе рождения приведена в открытых источниках'}${f.born ? ' в ' + f.born : ''}. ${f.country ? 'Гражданство: ' + f.country + '. ' : ''}${f.job ? 'Род деятельности: ' + dot(f.job) : ''}`);
  lines.push('', 'Основная часть');
  let ss = sentences(d.text || '', 8);
  if (parsed && parsed.noWar) ss = ss.filter(s => !/войн|боев|военн/i.test(s));
  if (parsed && parsed.noPolitics) ss = ss.filter(s => !/полит|парти|президент|правительств/i.test(s));
  if (ss.length) for (const s of ss) lines.push(s);
  else lines.push('В открытых источниках приведены основные сведения о жизни и деятельности этой личности.');

  if (f.education || f.career || (parsed && parsed.onlySchool)) {
    lines.push('', 'Образование и путь');
    if (f.education) lines.push('Образование: ' + dot(String(f.education).replace(/^.*?(учился|училась|окончил|окончила|образование)\s*/i, '')));
    if (f.career) lines.push('Деятельность: ' + dot(String(f.career).replace(/^.*?(деятельность|работал|работала|занимался|занималась)\s*/i, '')));
    if (!f.education && !f.career) lines.push('Подробные сведения об образовании и карьере приведены в открытых источниках.');
  }

  if (parsed && parsed.onlyRaised) {
    lines.push('', 'Воспитание');
    lines.push(f.raised ? dot(cap(f.raised)) : 'Сведения о воспитании приведены в открытых источниках.');
    const parents = Object.entries(f.family || {}).filter(([k]) => /отец|мать|родител/i.test(k));
    for (const [k, v] of parents) lines.push(`${k}: ${v.join(', ')}.`);
  }

  const fam = Object.entries(f.family || {});
  if (fam.length && !(parsed && parsed.noFamily)) {
    lines.push('', 'Семья');
    for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}.`);
  }

  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);

  let text = lines.join('\n');
  text = text.replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\s+([.,;:])/g, '$1')
    .trim();
  return text;
}

function sentences(text, limit = 6) {
  return String(text || '').replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ============================ PRESENTATION ============================

const THEMES = [
  { name: 'синий', bg: 'F2F6FF', dark: '1E3A8A', accent: '3B82F6', light: 'DBEAFE' },
  { name: 'зелёный', bg: 'F1F8F2', dark: '1B5E20', accent: '43A047', light: 'DCEDC8' },
  { name: 'фиолетовый', bg: 'F7F3FF', dark: '4A148C', accent: '7B1FA2', light: 'E1BEE7' },
  { name: 'granat', bg: 'F4F6FA', dark: '0F2A5A', accent: '2E5CA8', light: 'D6E4FF' },
  { name: 'тёплый', bg: 'FFF8EE', dark: '8A3B12', accent: 'D97706', light: 'FFE0B2' }
];

function pickTheme(title) {
  const h = [...norm(title)].reduce((a, c) => a + c.charCodeAt(0), 0);
  return THEMES[h % THEMES.length];
}

function splitBullets(text, max = 5) {
  const ss = sentences(text, max);
  const out = [];
  for (const s of ss) {
    let clean = s.replace(/^.*?(?:родился|родилась)[^,]{0,60},\s*/i, '')
      .replace(/^В\s+\d{4}\s+году,\s*/i, '');
    if (clean.length > 180) clean = clean.slice(0, 177).trim() + '…';
    out.push(clean || s);
  }
  return out.slice(0, max);
}

function buildSlidePlan(d, slideCount, parsed) {
  const n = Math.max(3, Math.min(16, Number(slideCount) || 7));
  const theme = pickTheme(d.title);
  const f = d.facts || {};

  const plan = [];
  plan.push({
    kind: 'title',
    title: d.title,
    subtitle: 'Доклад / презентация',
    badge: `${parsed.grade ? parsed.grade + ' класс' : 'Школьный проект'}`
  });

  plan.push({
    kind: 'bullets',
    title: 'Кто это',
    bullets: [
      `${d.title} — ${f.job || 'известная личность'}.`,
      f.birth ? `Год рождения: ${f.birth}.` : 'Сведения о жизни приведены в открытых источниках.',
      f.born ? `Место рождения: ${f.born}.` : '',
      f.country ? `Страна: ${f.country}.` : ''
    ].filter(Boolean)
  });

  const bio = splitBullets(d.text || '', 5);
  const bioChunks = [];
  for (let i = 0; i < bio.length; i += 3) bioChunks.push(bio.slice(i, i + 3));
  if (!bioChunks.length) bioChunks.push(['Основные сведения приведены в открытых источниках.']);

  for (const chunk of bioChunks) {
    plan.push({ kind: 'bullets', title: 'Биография', bullets: chunk });
  }

  if (f.education || f.career || parsed.onlySchool) {
    plan.push({
      kind: 'bullets',
      title: 'Образование и путь',
      bullets: [
        f.education ? 'Где учился(ась): ' + String(f.education).replace(/^.*?(учился|училась|окончил|окончила|образование)\s*/i, '') : 'Сведения об образовании приведены в открытых источниках.',
        f.career ? 'На кого пошёл(шла): ' + String(f.career).replace(/^.*?(деятельность|работал|работала|занимался|занималась)\s*/i, '') : 'После учёбы связал(а) жизнь с общественной деятельностью.'
      ].filter(Boolean)
    });
  }

  if (f.raised || parsed.onlyRaised) {
    plan.push({
      kind: 'bullets',
      title: 'Детство и воспитание',
      bullets: [
        f.raised || 'Воспитывался(ась) в семье, которая повлияла на его/её взгляды и интересы.',
        f.family && f.family['Отец'] ? 'Отец: ' + f.family['Отец'].join(', ') : '',
        f.family && f.family['Мать'] ? 'Мать: ' + f.family['Мать'].join(', ') : ''
      ].filter(Boolean)
    });
  }

  const fam = Object.entries(f.family || {});
  if (fam.length && !parsed.noFamily) {
    plan.push({
      kind: 'bullets',
      title: 'Семья',
      bullets: fam.slice(0, 4).map(([k, v]) => `${k}: ${v.slice(0, 2).join(', ')}`)
    });
  }

  plan.push({
    kind: 'bullets',
    title: 'Значение',
    bullets: [
      `${d.title} оставил(а) заметный след в истории.`,
      'Его/её жизнь и идеи изучают в школе.',
      'Биография подходит для доклада и самостоятельной работы.'
    ].filter(Boolean)
  });

  plan.push({
    kind: 'final',
    title: 'Вывод',
    bullets: [
      `${d.title} — личность, чья жизнь связана с важными событиями и идеями.`,
      parsed.noWar ? 'В докладе не рассматривались военные действия.' : '',
      parsed.noFamily ? 'Семья не рассматривалась по заданию.' : ''
    ].filter(Boolean)
  });

  // Fill or trim to requested count while preserving logical order.
  while (plan.length < n) {
    const idx = plan.length;
    const chunk = bio[(idx - 3) % Math.max(1, bio.length)] || 'Дополнительные сведения можно найти в открытых источниках.';
    plan.splice(Math.max(3, plan.length - 2), 0, { kind: 'bullets', title: 'Подробности', bullets: [chunk] });
  }
  if (plan.length > n) {
    const fixed = plan.filter(p => p.kind !== 'bullets').length;
    const keep = Math.max(0, n - fixed);
    const bullets = plan.filter(p => p.kind === 'bullets');
    const other = plan.filter(p => p.kind !== 'bullets');
    const rebuilt = [];
    let bi = 0;
    for (const p of plan) {
      if (p.kind === 'bullets') {
        if (bi++ < keep) rebuilt.push(p);
      } else rebuilt.push(p);
    }
    return rebuilt.slice(0, n);
  }
  return plan;
}

async function makePptx(d, slideCount, parsed) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'SchoolBot';
  pptx.subject = 'Школьный доклад';
  pptx.title = d.title;
  pptx.lang = 'ru-RU';
  pptx.theme = {
    headFontFace: 'Arial', bodyFontFace: 'Arial', lang: 'ru-RU'
  };

  const plan = buildSlidePlan(d, slideCount, parsed);
  const theme = pickTheme(d.title);

  for (const item of plan) {
    const slide = pptx.addSlide();
    slide.background = { color: theme.bg };

    // Top colored band (slide size is 10 x 5.625 in).
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: '100%', h: 0.12, fill: { color: theme.accent }, line: { color: theme.accent }
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0.12, w: '100%', h: 0.05, fill: { color: theme.light }, line: { color: theme.light }
    });

    if (item.kind === 'title') {
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.5, y: 0.5, w: '90%', h: 1.0,
        fill: { color: theme.dark }, line: { color: theme.dark }, radius: 0.12
      });
      slide.addText(item.title, {
        x: 0.65, y: 0.63, w: '85%', h: 0.75, align: 'center', valign: 'mid',
        fontSize: 30, bold: true, color: 'FFFFFF', fontFace: 'Arial', margin: 0
      });
      if (item.subtitle) slide.addText(item.subtitle, {
        x: 0, y: 4.35, w: '100%', h: 0.35, align: 'center', fontSize: 16, color: theme.dark
      });
      slide.addText(item.badge || '', {
        x: 0, y: 4.85, w: '100%', h: 0.35, align: 'center', fontSize: 14,
        color: theme.dark, fontFace: 'Arial'
      });
      continue;
    }

    // Regular slide: centered top heading.
    slide.addText(item.title, {
      x: 0.5, y: 0.3, w: '90%', h: 0.6,
      align: 'center', valign: 'mid', fontSize: 24, bold: true,
      color: theme.dark, fontFace: 'Arial', margin: 0
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 2.0, y: 0.95, w: 6.0, h: 0.04, fill: { color: theme.accent }, line: { color: theme.accent }
    });

    const bullets = (item.bullets || []).filter(Boolean);
    const fontSize = bullets.length <= 2 ? 24 : bullets.length <= 4 ? 21 : 18;
    const bodyY = 1.15;
    const bodyH = bullets.length <= 2 ? 3.5 : 4.15;

    slide.addText(
      bullets.map((b, i) => ({ text: String(b).replace(/\s+/g, ' ').trim(), options: { bullet: { code: '2022' }, breakLine: i < bullets.length - 1 } })),
      {
        x: 0.7, y: bodyY, w: '86%', h: bodyH, valign: 'mid', align: 'center',
        fontSize, color: '1F2937', fontFace: 'Arial', lineSpacingMultiple: 1.18,
        paraSpaceAfterPt: 12, margin: 0.08, fit: 'shrink'
      }
    );

    // Decorative accent element for a modern, non-plain look.
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 9.2, y: 5.05, w: 0.25, h: 0.25, fill: { color: theme.accent, transparency: 40 }, line: { color: theme.accent }
    });
  }

  const outDir = path.join(__dirname, 'generated');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const file = path.join(outDir, `presentation_${Date.now()}.pptx`);
  await pptx.writeFile({ fileName: file });
  return file;
}

// ============================ TELEGRAM ============================

const bot = new Telegraf(TOKEN);

const mainMenu = Markup.keyboard([
  ['📝 Готовый текст', '📊 Презентация']
]).resize().reply_markup;

const removeKb = Markup.removeKeyboard().reply_markup;

const sessions = new Map();
function getSession(id) {
  if (!sessions.has(id)) sessions.set(id, { step: 'menu', awaitingPres: false, name: '', note: '' });
  return sessions.get(id);
}

async function safeReply(ctx, text, markup) {
  const chunks = String(text).match(/[\s\S]{1,3800}/g) || [text];
  for (const chunk of chunks) {
    try { await ctx.reply(chunk, markup ? { reply_markup: markup } : undefined); }
    catch (e) { console.error('reply error', e.message); }
  }
}

async function handleTextRequest(ctx, s, raw) {
  const parsed = parseRequest(raw, 'text');
  await safeReply(ctx, 'Понял запрос. Ищу и сверяю по нескольким источникам (до 30 секунд)…');

  let person = null;
  try {
    person = await withTimeout(collectPerson(raw, parsed), MAX_SEARCH_MS, 'search');
  } catch (e) {
    console.error('search error', e.message);
  }

  if (!person) {
    await safeReply(ctx, 'Не смог понять, о ком речь. Напиши, пожалуйста, имя человека целиком, например: «Альберт Эйнштейн» или «Уинстон Черчилль».');
    return;
  }

  const report = makeReport(person, raw, parsed);
  s.name = person.title;
  s.note = raw;
  s.step = 'menu';
  await safeReply(ctx, report);
  await safeReply(ctx, `Источник: ${person.source}. Если нужно изменить доклад, напиши уточнение (например, «без войны», «8 класс», «где учился»).`, mainMenu);
}

async function handlePresRequest(ctx, s, raw) {
  const parsed = parseRequest(raw, 'pres');
  s.storedParsed = parsed;

  const nums = String(raw).match(/\b(\d{1,2})\s*(?:слайд|слайдов|слайда|лист|листов)\b/i) ||
    String(raw).match(/\b(\d{1,2})\s*[.)]?\s*$/);
  s.slideCount = nums ? Math.max(3, Math.min(16, Number(nums[1]))) : 0;

  if (!s.slideCount) {
    s.step = 'pres_slides';
    await safeReply(ctx, 'Сколько слайдов нужно? Напиши числом от 3 до 16 (например, 8).', removeKb);
    return;
  }

  await generatePres(ctx, s, raw);
}

async function generatePres(ctx, s, raw) {
  const parsed = s.storedParsed || parseRequest(raw || s.name || '', 'pres');
  await safeReply(ctx, 'Готовлю красивую презентацию…');

  let person = null;
  try {
    const query = s.name || raw || '';
    person = await withTimeout(collectPerson(query, parsed), MAX_SEARCH_MS, 'search');
  } catch (_) {}

  if (!person) {
    await safeReply(ctx, 'Не смог понять, о ком презентация. Напиши имя целиком.', mainMenu);
    s.step = 'menu';
    return;
  }

  try {
    const file = await makePptx(person, s.slideCount, parsed);
    await ctx.replyWithDocument({ source: file, filename: `Презентация_${person.title.replace(/[^\wа-яА-ЯёЁ-]+/g, '_')}.pptx` });
    await safeReply(ctx, `Готово! Слайдов: ${s.slideCount}. Оформление подобрано автоматически.`, mainMenu);
  } catch (e) {
    console.error('pptx error', e);
    await safeReply(ctx, 'Не удалось собрать файл презентации. Попробуй ещё раз.', mainMenu);
  }
  s.step = 'menu';
}

bot.on('text', async (ctx, next) => {
  if (!ctx.message || !ctx.message.text) return next();
  const text = ctx.message.text.trim();
  const s = getSession(ctx.from.id);

  if (text === '/start') {
    s.step = 'menu';
    await safeReply(ctx, `Привет! Я помогаю делать школьные доклады и презентации о публичных личностях. Версия ${BOT_VERSION}.\nВыбери кнопку или просто напиши, про кого нужно.`, mainMenu);
    return;
  }

  if (text === '📝 Готовый текст') {
    s.step = 'text_topic';
    await safeReply(ctx, 'Напиши, про кого и что учесть. Например:\nУинстон Черчилль, 8 класс, без войны, где учился и на кого пошёл', removeKb);
    return;
  }

  if (text === '📊 Презентация') {
    s.step = 'pres_topic';
    await safeReply(ctx, 'Напиши, про кого презентация и сколько слайдов. Например:\nУинстон Черчилль, 8 слайдов', removeKb);
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    s.name = text;
    await handlePresRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_slides') {
    await generatePres(ctx, s, text);
    return;
  }

  // Free-form message from menu: if it looks like a presentation request, make pptx; else report.
  if (/презентац|слайд|презент/i.test(text)) {
    s.name = text;
    await handlePresRequest(ctx, s, text);
  } else {
    await handleTextRequest(ctx, s, text);
  }
});

bot.on('document', ctx => ctx.reply('Пришли, пожалуйста, текстом имя человека, про которого нужен доклад.'));

// ============================ START ============================
'use strict';

// School report bot. Public figures and public data only.
// Local mode: node school_bot.js (uses Happ proxy if available).
// Cloud mode: set BOT_TOKEN and WEBHOOK_URL; app listens on PORT.

const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { Telegraf, Markup } = require('telegraf');
const PptxGenJS = require('pptxgenjs');
const express = require('express');
const HttpsProxyAgentClass = require('https-proxy-agent');

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) throw new Error('BOT_TOKEN is not set');
const USE_PROXY = !!process.env.USE_PROXY;
const PROXY_URL = process.env.PROXY_URL || 'socks://127.0.0.1:10808';
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.WEBHOOK_URL || '';
const BOT_VERSION = '2026-11-21-v4';

const WIKI_API = 'https://ru.wikipedia.org/w/api.php';
const WIKI_FALLBACK = 'https://ru.m.wikipedia.org/w/api.php';
const WIKIDATA_API = 'https://www.wikidata.org/w/api.php';
const SPARQL_API = 'https://query.wikidata.org/sparql';
const MAX_SEARCH_MS = 29000;
const MAX_REQUEST_MS = 9000;

const http = axios.create({
  timeout: MAX_REQUEST_MS,
  headers: {
    'User-Agent': 'SchoolBot/1.0 (school project; t.me/schol_spg_bot)',
    Accept: 'application/json,text/plain,*/*'
  },
  proxy: false
});
if (USE_PROXY) http.defaults.httpsAgent = new HttpsProxyAgentClass(PROXY_URL);

// ============================ DATA ============================

// Canonical public figures with manually checked, school-safe facts.
const LOCAL_PEOPLE = [
  {
    name: 'Уинстон Черчилль',
    aliases: [
      'чертиль', 'чертиля', 'чертил', 'чертильч', 'черчиль', 'черчильч',
      'черчилль', 'черчилля', 'уинстон чертиль', 'уинстона чертиля',
      'уинстон черчилль', 'уинстона черчилля', 'уинстон черчиль',
      'увинстона чертиля', 'увинстон чертиль', 'wineston churchill',
      'winston churchill', 'уинстон churchill'
    ],
    facts: {
      birth: 1874, death: 1965, born: 'Бленхейм, Англия', country: 'Великобритания',
      job: 'политик, премьер-министр, писатель',
      education: 'Королевская военная академия в Сандхёрсте',
      career: 'офицер, журналист, политик; премьер-министр Великобритании в 1940–1945 и 1951–1955 гг.',
      raised: 'в аристократической семье; отец — лорд Рандольф Черчилль, мать — Дженни Джером',
      family: { 'Отец': ['Рандольф Черчилль'], 'Мать': ['Дженни Джером'], 'Жена': ['Клементина Черчилль'] }
    },
    text: 'Уинстон Черчилль родился 30 ноября 1874 года в Бленхеймском дворце. Учился в Королевской военной академии в Сандхёрсте, затем служил офицером и работал военным корреспондентом. После этого он пошёл в политику: неоднократно избирался в парламент, занимал министерские посты, а в 1940 году стал премьер-министром Великобритании. В годы Второй мировой войны его речи и решения пользовались большой поддержкой в стране. После войны он снова вернулся в политику, а также написал много исторических и мемуарных книг. В 1953 году Черчилль получил Нобелевскую премию по литературе.'
  },
  {
    name: 'Альберт Эйнштейн',
    aliases: ['энштейн', 'энштейна', 'ейнштейн', 'айнштейн', 'einstein', 'альберт энштейн', 'альберта энштейна'],
    facts: {
      birth: 1879, death: 1955, born: 'Ульм, Германия', country: 'Германия / США',
      job: 'физик-теоретик',
      education: 'школа в Мюнхене, Цюрихский политехникум',
      career: 'работал в патентном бюро, затем профессором и исследователем; создал теорию относительности',
      raised: 'в семье инженера и предпринимателя; дядя поощрял интерес к науке',
      family: { 'Отец': ['Герман Эйнштейн'], 'Мать': ['Паулина Кох'], 'Жена': ['Милева Марич', 'Эльза Левенталь'] }
    },
    text: 'Альберт Эйнштейн родился 14 марта 1879 года в Ульме. Детство прошло в Мюнхене, где его семья владела фирмой по производству электрооборудования. После переезда в Италию он продолжил обучение в Швейцарии и окончил Цюрихский политехникум. Некоторое время работал в патентном бюро, а затем занялся научной деятельностью. В 1905 году он опубликовал работы по специальной теории относительности, фотоэффекту и броуновскому движению. Позже создал общую теорию относительности. В 1921 году получил Нобелевскую премию по физике. В 1930-х годах переехал в США и стал одним из самых известных учёных XX века.'
  },
  {
    name: 'Александр Пушкин',
    aliases: ['пушкин', 'пушкина', 'александр пушкин', 'александра пушкина', 'саша пушкин'],
    facts: {
      birth: 1799, death: 1837, born: 'Москва', country: 'Россия',
      job: 'поэт, прозаик, драматург',
      education: 'Лицей в Царском Селе',
      career: ' начинал как поэт, служил в Министерстве иностранных дел, затем полностью посвятил себя литературе',
      raised: 'в дворянской семье; воспитанием и образованием занимались гувернёры и наставники лицея',
      family: { 'Отец': ['Сергей Пушкин'], 'Мать': ['Надежда Ганнибал'], 'Жена': ['Наталья Гончарова'] }
    },
    text: 'Александр Сергеевич Пушкин родился 26 мая 1799 года в Москве. Отец происходил из старинного дворянского рода, мать была внучкой африканца Ибрагима Ганнибала. В 1811 году Пушкин поступил в только что открытый Лицей в Царском Селе, где получил отличное образование. После окончания лицея он служил в Министерстве иностранных дел, но главным делом жизни стала литература. Пушкин создал роман «Евгений Онегин», трагедию «Борис Годунов», «Капитанскую дочку», много стихотворений и сказок. Он считается основоположником современной русской литературной формы.'
  },
  {
    name: 'Юрий Гагарин',
    aliases: ['гагарин', 'гагарина', 'юра гагарин', 'первый космонавт'],
    facts: {
      birth: 1934, death: 1968, born: 'Клушино, Смоленская область', country: 'СССР',
      job: 'космонавт, лётчик',
      education: 'Саратовский индустриальный техникум, Оренбургское военное авиационное училище',
      career: ' стал лётчиком, затем прошёл подготовку в отряде космонавтов; 12 апреля 1961 года совершил первый в истории полёт в космос',
      raised: 'в семье колхозников; после войны семья переехала в Гжатск',
      family: { 'Отец': ['Алексей Гагарин'], 'Мать': ['Анна Гагарина'], 'Жена': ['Валентина Гагарина'] }
    },
    text: 'Юрий Алексеевич Гагарин родился 9 марта 1934 года в деревне Клушино. Учился в Саратовском индустриальном техникуме, затем в Оренбургском авиационном училище. Пошёл по стопам старшего брата и стал лётчиком. В 1960 году был зачислен в первый отряд космонавтов. 12 апреля 1961 года на корабле «Восток-1» он впервые в истории человека совершил орбитальный полёт вокруг Земли. После полёта стал всемирно известным, занимался общественной деятельностью и подготовкой новых космонавтов.'
  },
  {
    name: 'Лев Толстой',
    aliases: ['толстой', 'толстого', 'лев толстой', 'льва толстого', 'граф толстой'],
    facts: {
      birth: 1828, death: 1910, born: 'Ясная Поляна, Тульская губерния', country: 'Россия',
      job: 'писатель, мыслитель',
      education: 'домашнее образование, затем Казанский университет',
      career: 'участвовал в Крымской войне, затем занимался литературой и педагогикой',
      raised: 'в дворянской семье; получил домашнее воспитание',
      family: { 'Отец': ['Николай Толстой'], 'Мать': ['Мария Волконская'], 'Жена': ['София Толстая'] }
    },
    text: 'Лев Николаевич Толстой родился в 1828 году в Ясной Поляне. Получил домашнее образование, затем учился в Казанском университете. После военной службы занялся литературой. Главные произведения — романы «Война и мир», «Анна Каренина», «Воскресение». В последние годы жизни Толстой отказался от дворянского образа жизни, проповедовал ненасилие и простоту. Его творчество оказало огромное влияние на русскую и мировую литературу.'
  },
  {
    name: 'Наполеон Бонапарт',
    aliases: ['наполеон', 'наполеона', 'бонапарт', 'бонапарта', 'наполеон бонапарт', 'наполеона бонапарта'],
    facts: {
      birth: 1769, death: 1821, born: 'Аяччо, Корсика', country: 'Франция',
      job: 'полководец, император',
      education: 'военная школа в Бриенне, Парижская военная школа',
      career: 'военный офицер, затем консул и император Франции',
      raised: 'в небогатой дворянской семье на Корсике; воспитывался в военных школах',
      family: { 'Отец': ['Карло Бонапарт'], 'Мать': ['Летиция Рамолино'], 'Жена': ['Жозефина', 'Мария Луиза'] }
    },
    text: 'Наполеон Бонапарт родился в 1769 году в Аяччо на Корсике. Учился в военных школах во Франции и рано начал военную службу. Во время Французской революции быстро выдвинулся, затем стал консулом, а в 1804 году — императором. Он провёл много военных кампаний и создал крупную империю, но после поражения в 1814 году отрёкся от престола. В 1815 году попытался вернуться к власти, но потерпел поражение при Ватерлоо и был сослан на остров Святой Елены.'
  },
  {
    name: 'Мария Кюри',
    aliases: ['мария кюри', 'марию кюри', 'кюри', 'мари курi', 'curie', 'мария склодовская-кюри'],
    facts: {
      birth: 1867, death: 1934, born: 'Варшава', country: 'Польша / Франция',
      job: 'физик и химик',
      education: 'Сорбонна в Париже',
      career: 'исследователь радиоактивности, профессор Сорбонны',
      raised: 'в семье учителей; с детства проявляла интерес к естественным наукам',
      family: { 'Отец': ['Владислав Склодовский'], 'Мать': ['Бронислава Склодовская'], 'Муж': ['Пьер Кюри'] }
    },
    text: 'Мария Склодовская-Кюри родилась в 1867 году в Варшаве. Училась в подпольных польских школах, затем в Сорбонне в Париже. вместе с мужем Пьером Кюри исследовала радиоактивные элементы и открыла полоний и радий. Дважды получала Нобелевскую премию: по физике в 1903 году и по химии в 1911 году. Она стала первой женщиной — профессором Сорбонны.'
  },
  {
    name: 'Михаил Ломоносов',
    aliases: ['ломоносов', 'ломоносова', 'михаил ломоносов', 'михаила ломоносова'],
    facts: {
      birth: 1711, death: 1765, born: 'Денисовка, Архангельская губерния', country: 'Россия',
      job: 'учёный, поэт',
      education: 'Славяно-греко-латинская академия в Москве, университет в Марбурге',
      career: 'профессор химии, основатель Московского университета',
      raised: 'в семье поморского крестьянина; с юности стремился к учёбе',
      family: { 'Отец': ['Василий Ломоносов'], 'Мать': ['Елена Ломоносова'], 'Жена': ['Елизавета Дильбах'] }
    },
    text: 'Михаил Васильевич Ломоносов родился в 1711 году в поморской деревне. В 19 лет пешком отправился в Москву и поступил в Славяно-греко-латинскую академию. Затем учился в Германии. Вернувшись в Россию, стал профессором, занимался химией, физикой, геологией и поэзией. Вместе с Шуваловым основал Московский университет. Его работы заложили основы русской научной традиции.'
  },
  {
    name: 'Иван Грозный',
    aliases: ['иван грозный', 'ивана грозного', 'грозный', 'иван iv', 'иван васильевич'],
    facts: {
      birth: 1530, death: 1584, born: 'Москва', country: 'Русское царство',
      job: 'царь',
      education: 'домашнее образование при дворе',
      career: 'великий князь московский, затем первый русский царь',
      raised: 'при дворе в окружении бояр; детство прошло в условиях придворных конфликтов',
      family: { 'Отец': ['Василий III'], 'Мать': ['Елена Глинская'], 'Жены': ['Анастасия Романова', 'Мария Темрюковна и др.'] }
    },
    text: 'Иван IV Васильевич родился в 1530 году. В 1547 году был коронован как первый русский царь. При нём были проведены реформы, созданы приказы, начата Ливонская война. В памяти народа остался как жестокий правитель эпохи опричнины, но также как создатель централизованного Русского государства.'
  },
  {
    name: 'Пётр I',
    aliases: ['пётр первый', 'петр первый', 'пётр 1', 'петр 1', 'пётр великий', 'петр великий', 'пётр', 'петр алексеевич'],
    facts: {
      birth: 1672, death: 1725, born: 'Москва', country: 'Россия',
      job: 'царь, реформатор',
      education: 'домашнее образование, затем обучение в Голландии и Англии',
      career: 'царь, создал регулярную армию и флот, основал Санкт-Петербург',
      raised: 'в царской семье; с детства интересовался военным делом и мореплаванием',
      family: { 'Отец': ['Алексей Михайлович'], 'Мать': ['Наталья Нарышкина'], 'Жена': ['Екатерина I'] }
    },
    text: 'Пётр I родился в 1672 году. С юности увлекался военным делом и кораблями. В 1690-х годах его называли «царём-плотником». Он провёл реформы армии, управления и образования, основал Санкт-Петербург и Академию наук. После победы в Северной войне Россия вышла к Балтийскому морю. Пётр I считается одним из главных реформаторов в русской истории.'
  },
  {
    name: 'Карл Маркс',
    aliases: ['маркс', 'маркса', 'карл маркс', 'карла маркса'],
    facts: {
      birth: 1818, death: 1883, born: 'Трир, Германия', country: 'Германия',
      job: 'философ, экономист',
      education: 'университеты Бонна и Берлина',
      career: 'журналист и теоретик; создал учение о капитализме вместе с Энгельсом',
      raised: 'в семье адвоката; получил классическое образование',
      family: { 'Отец': ['Генрих Маркс'], 'Мать': ['Генриетта Прево'], 'Жена': ['Женни фон Вестфален'] }
    },
    text: 'Карл Маркс родился в 1818 году в Трире. Учился в университетах Бонна и Берлина. Занимался журналистикой, но из-за своих взглядов не смог работать в Германии и переехал в Париж, затем в Лондон. Вместе с Фридрихом Энгельсом написал «Манифест Коммунистической партии» и «Капитал». Его идеи оказали огромное влияние на историю XX века.'
  },
  {
    name: 'Фридрих Энгельс',
    aliases: ['энгельс', 'энгельса', 'фридрих энгельс', 'фридриха энгельса'],
    facts: {
      birth: 1820, death: 1895, born: 'Бармен, Германия', country: 'Германия',
      job: 'философ, экономист',
      education: 'гимназия, коммерческое училище',
      career: 'бизнесмен, журналист и теоретик; соавтор Маркса',
      raised: 'в семье фабриканта',
      family: { 'Отец': ['Фридрих Энгельс-старший'], 'Мать': ['Элизабет ван Хорн'] }
    },
    text: 'Фридрих Энгельс родился в 1820 году. Работал в торговой фирме в Манчестере и изучал положение рабочих. Вместе с Карлом Марксом написал «Манифест Коммунистической партии» и другие работы. После смерти Маркса подготовил к печати второй и третий тома «Капитала».'
  },
  {
    name: 'Уильям Шекспир',
    aliases: ['шекспир', 'шекспира', 'уильям шекспир', 'вильям шекспир', 'shakespeare'],
    facts: {
      birth: 1564, death: 1616, born: 'Стратфорд-на-Эйвоне', country: 'Англия',
      job: 'драматург, поэт',
      education: 'школа в Стратфорде',
      career: 'актёр, драматург, владелец театра «Глобус»',
      raised: 'в семье перчаточника и мелкого чиновника',
      family: { 'Отец': ['Джон Шекспир'], 'Мать': ['Мэри Арден'], 'Жена': ['Анна Хэтэуэй'] }
    },
    text: 'Уильям Шекспир родился в 1564 году в Стратфорде. Учился в местной школе, затем стал актёром и драматургом в Лондоне. Написал 37 пьес, среди них «Гамлет», «Макбет», «Король Лир», «Ромео и Джульетта». Его произведения переведены на многие языки и до сих пор ставятся в театрах.'
  },
  {
    name: 'Чарльз Дарвин',
    aliases: ['дарвин', 'дарвина', 'чарльз дарвин', 'чарлз дарвин', 'darwin'],
    facts: {
      birth: 1809, death: 1882, born: 'Шрусбери, Англия', country: 'Великобритания',
      job: 'естествоиспытатель',
      education: 'школа в Шрусбери, Эдинбургский и Кембриджский университеты',
      career: 'путешественник и исследователь; создал теорию эволюции',
      raised: 'в семье врача; с детства любил собирать коллекции',
      family: { 'Отец': ['Роберт Дарвин'], 'Мать': ['Сюзанна Уэджвуд'], 'Жена': ['Эмма Уэджвуд'] }
    },
    text: 'Чарльз Дарвин родился в 1809 году. Учился в школе, затем в университетах. В 1831 году отправился в плавание на корабле «Бигль», что дало ему материал для наблюдений. В 1859 году опубликовал книгу «Происхождение видов», где объяснил теорию эволюции естественным отбором. Его идеи изменили биологию.'
  },
  {
    name: 'Александр Невский',
    aliases: ['александр невский', 'александра невского', 'невский', 'невского'],
    facts: {
      birth: 1220, death: 1263, born: 'Переславль-Залесский', country: 'Древняя Русь',
      job: 'князь, полководец',
      education: 'домашнее образование',
      career: 'новгородский князь; победил шведов на Неве и крестоносцев на Чудском озере',
      raised: 'в семье князя Ярослава Всеволодовича',
      family: { 'Отец': ['Ярослав Всеволодович'], 'Мать': ['Феодосия'], 'Жена': ['Васса'] }
    },
    text: 'Александр Невский родился в 1220 году. Был новгородским князем. В 1240 году разбил шведское войско на реке Неве, а в 1242 году — рыцарей на Чудском озере. При нём Новгород сохранил независимость, а сам он стал символом защиты русских земель. Признан святым Русской православной церковью.'
  },
  {
    name: 'Дмитрий Менделеев',
    aliases: ['менделеев', 'менделеева', 'дмитрий менделеев', 'дмитрия менделеева'],
    facts: {
      birth: 1834, death: 1907, born: 'Тобольск', country: 'Россия',
      job: 'химик',
      education: 'Тобольская гимназия, Петербургский университет',
      career: 'профессор химии; открыл периодический закон и создал таблицу элементов',
      raised: 'в большой семье директора гимназии',
      family: { 'Отец': ['Иван Менделеев'], 'Мать': ['Мария Корнилова'], 'Жена': ['Анна Попова'] }
    },
    text: 'Дмитрий Иванович Менделеев родился в 1834 году в Тобольске. Учился в гимназии и Петербургском университете. В 1869 году открыл периодический закон химических элементов и составил периодическую таблицу. Также занимался вопросами промышленности и метрологии.'
  },
  {
    name: 'Исаак Ньютон',
    aliases: ['ньютон', 'ньютона', 'исаак ньютон', 'исаака ньютона', 'newton'],
    facts: {
      birth: 1643, death: 1727, born: 'Вулсторп, Англия', country: 'Великобритания',
      job: 'физик, математик',
      education: 'Тринити-колледж Кембриджского университета',
      career: 'профессор математики; создал классическую механику и теорию тяготения',
      raised: 'в семье фермера; воспитывался бабушкой после смерти матери',
      family: { 'Отец': ['Исаак Ньютон-старший'], 'Мать': ['Ханна Эйскью'] }
    },
    text: 'Исаак Ньютон родился в 1643 году. Учился в Кембридже. Сформулировал законы движения и закон всемирного тяготения, создал основы дифференциального исчисления. Его книга «Математические начала натуральной философии» стала одним из главных научных трудов в истории.'
  },
  {
    name: 'София Ковалевская',
    aliases: ['ковалевская', 'ковалевской', 'софия ковалевская', 'софья ковалевская', 'софии ковалевской'],
    facts: {
      birth: 1850, death: 1891, born: 'Москва', country: 'Россия / Швеция',
      job: 'математик',
      education: 'Гейдельбергский и Берлинский университеты',
      career: 'первая женщина — профессор математики в Северной Европе',
      raised: 'в дворянской семье; интерес к математике проявился в детстве',
      family: { 'Отец': ['Василий Корвин-Круковский'], 'Мать': ['Елизавета Шуберт'], 'Муж': ['Владимир Ковалевский'] }
    },
    text: 'София Ковалевская родилась в 1850 году в Москве. Из-за запрета женщинам учиться в российских университетах она уехала в Германию и слушала лекции в Гейдельберге и Берлине. Стала первой женщиной, получившей докторскую степень по математике, и первой женщиной-профессором в Стокгольмском университете.'
  },
  {
    name: 'Владимир Ленин',
    aliases: ['ленин', 'ленина', 'владимир ленин', 'владимира ленина', 'ульянов', 'ульянова'],
    facts: {
      birth: 1870, death: 1924, born: 'Симбирск', country: 'Россия / СССР',
      job: 'политик, революционер',
      education: 'Симбирская гимназия, Казанский университет',
      career: 'революционер, основатель большевистской партии, первый глава советского государства',
      raised: 'в семье учителя; старший брат был казнён, что повлияло на его взгляды',
      family: { 'Отец': ['Илья Ульянов'], 'Мать': ['Мария Бланк'], 'Жена': ['Надежда Крупская'] }
    },
    text: 'Владимир Ильич Ленин родился в 1870 году в Симбирске. Учился в гимназии и Казанском университете. После казни старшего брата увлёкся революционными идеями. В 1917 году возглавил большевиков и Октябрьскую революцию. С 1917 года возглавлял советское правительство.'
  },
  {
    name: 'Александр Суворов',
    aliases: ['суворов', 'суворова', 'александр суворов', 'александра суворова'],
    facts: {
      birth: 1730, death: 1800, born: 'Москва', country: 'Россия',
      job: 'полководец',
      education: 'домашнее образование',
      career: 'военный командир; провёл Итальянский и Швейцарский походы',
      raised: 'в семье военного; с детства готовился к службе',
      family: { 'Отец': ['Василий Суворов'], 'Мать': ['Мария Манукова'] }
    },
    text: 'Александр Васильевич Суворов родился в 1730 году. С детства готовился к военной службе. Отличился в Русско-турецких войнах, при взятии Измаила, в Итальянском и Швейцарском походах. Известен как выдающийся полководец, никогда не проигрывавший сражения.'
  }
];

const ALIASES = {};
for (const p of LOCAL_PEOPLE) {
  ALIASES[norm(p.name)] = p.name;
  for (const a of p.aliases) ALIASES[norm(a)] = p.name;
}

const STOP_WORDS = new Set([
  'привет','здравствуй','здравствуйте','хай','добрый','день','утро','вечер',
  'мне','надо','нужно','нжен','нужна','хочу','хочет','пожалуйста','пжлст','пж',
  'напиши','написать','сделай','сделать','сгенерируй','создай','давай','дай','пиши',
  'про','о','об','обо','вая','во','информация','инфу','инфо','сведения','расскажи',
  'рассказать','даные','данные',' доклад','реферат','сочинение','текст','письменная',
  'работа',' докладчик','презентация','презентации','презент','през','слайды','слайд',
  'сделай','сделать','на tema','тема','тему','это','этот','эта','эти','тот','та','того',
  'какой','какая','какие','кто','что','как','где','когда','почему','зачем','для',
  'школьный','школа',' clas','класс','класса','ученик','ученица','учитель',
  'хороший','красивый','красиво','подробно','подробная','подробный','коротко',
  'кратко','просто','попроще','серьёзно','серьезно','быстро','ещё','еще','и','а','но',
  'или','да','нет','не','без','без','же','бы','в','во','на','за','из','к','ко','у','от',
  'до','при','по','с','со','вот','уже','тоже','так','итак','например','например',
  'сам','сама','сами','себя','его','её','ее','их','його','им','ому','ему','ей','ихний',
  'нд','нде','н надо','ненадо','не надо','не надо','ненадо','ненад','ненада',
  'семью','семя','семьи','семей','семью','жену','мужа','детень','дети','детей',
  'войну','война','про войну','военный','военные','боевые','боевые действия',
  'учился','училась','где учился','где он учился','где она училась','образование',
  'пошёл','пошла','на кого пошёл','путь','карьера','деятельность','род деятельности',
  'воспитывал','воспитывала','кто воспитывал','родители','мама','папа','бабушка',
  'дедушка','род','родом','родился','родилась','биография','жизнь',' открытия',
  'вывод','заключение','введение','основная часть','глава','пункт','список',
  'можешь','сможешь',' есть','имеется','найди','найти','поищи','проверь','сверь',
  'источники','источник','Википедия','Вики','Так','У','КБЖД','готовый','готовое',
  'вариант','пример','шаблон','structure','структура','по пунктам','по пунктам'
]);

const RU_MAP = {
  а:['a','а'], б:['b','б'], в:['v','в'], г:['g','г'], д:['d','д'], е:['e','е'], ё:['e','ё'],
  ж:['zh','ж'], з:['z','з'], и:['i','и'], й:['y','й'], к:['k','к'], л:['l','л'], м:['m','м'],
  н:['n','н'], о:['o','о'], п:['p','п'], р:['r','р'], с:['s','с'], т:['t','т'], у:['u','у'],
  ф:['f','ф'], х:['h','х'], ц:['c','ц'], ч:['ch','ч'], ш:['sh','ш'], щ:['sch','щ'],
  ъ:['','ъ'], ы:['y','ы'], ь:['','ь'], э:['e','э'], ю:['u','ю'], я:['ya','я']
};
const TRANSLIT = {
  a:'а', b:'б', v:'в', g:'г', d:'д', e:'е', zh:'ж', z:'з', i:'и', y:'й', k:'к', l:'л',
  m:'м', n:'н', o:'о', p:'п', r:'р', s:'с', t:'т', u:'у', f:'ф', h:'х', c:'ц', ch:'ч',
  sh:'ш', sch:'щ', w:'в', x:'кс'
};

// ============================ UTIL ============================

function dot(s) {
  return String(s || '').replace(/\.+$/, '') + '.';
}

function cap(s) {
  const v = String(s || '').trim();
  return v ? v[0].toUpperCase() + v.slice(1) : v;
}

function norm(s) {
  return String(s || '').toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripWordEndings(token) {
  // Remove common Russian inflections for name search, keep enough for fuzzy match.
  return token
    .replace(/(ого|его|ому|ему|ыми|ими|ами|ями|ах|ях|ов|ев|ой|ей|ом|ем|у|ю|а|я|ы|и|е|ь)$/g, '');
}

function levenshtein(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const prev = new Array(b.length + 1);
  const cur = new Array(b.length + 1);
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    cur[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        cur[j] = Math.min(cur[j], prev[j - 2] + 1);
      }
    }
    for (let j = 0; j <= b.length; j++) prev[j] = cur[j];
  }
  return prev[b.length];
}

function t9Variants(word) {
  // If Latin typed in Russian layout or vice versa, transliterate both ways.
  const w = norm(word).replace(/\s/g, '');
  const cyr = /[а-я]/.test(w);
  let out = '';
  for (const ch of w) {
    if (/[a-z]/.test(ch)) {
      out += TRANSLIT[ch] || ch;
    } else if (RU_MAP[ch]) {
      out += RU_MAP[ch][0];
    } else out += ch;
  }
  return [w, out].filter(Boolean);
}

function similarity(a, b) {
  if (!a || !b) return 0;
  const max = Math.max(a.length, b.length);
  return Math.max(0, 1 - levenshtein(a, b) / max);
}

function bestNameScore(candidate, targetTokens) {
  const candTokens = norm(candidate).split(' ').filter(Boolean);
  let total = 0;
  let used = 0;
  for (const ct of candTokens) {
    const stem = stripWordEndings(ct);
    let best = 0;
    for (const tt of targetTokens) {
      for (const v of t9Variants(tt)) {
        const vStem = stripWordEndings(v);
        const s = Math.max(similarity(ct, v), similarity(stem, vStem), similarity(stem, v));
        if (s > best) best = s;
      }
    }
    total += best;
    if (best > 0.5) used++;
  }
  const coverage = used / Math.max(1, candTokens.length);
  const extraPenalty = Math.max(0, targetTokens.length - candTokens.length) * 0.08;
  return total / Math.max(1, candTokens.length) * 0.72 + coverage * 0.28 - extraPenalty;
}

// Parse all user constraints from the entire message.
function parseRequest(raw, mode) {
  const original = String(raw || '').trim();
  const lower = norm(original);
  const constraints = [];
  const add = (flag, phrase) => { if (flag && !constraints.includes(phrase)) constraints.push(phrase); };

  const noFamily = /без\s+сем(?:ью|ьи|и|я)|без\s+[а-яё]+\s+и\s+сем(?:ью|ьи|и|я)|сем(?:ью|ьи|и|я)\s*н[еи]?\s*над|н[еи]?\s*над[оа]?\s+сем(?:ью|ьи|и|я)|без\s+жен|без\s+муж|без\s+дет|сем(?:ью|ьи|и|я)\s+не\s+пиши|не\s+пиши\s+сем(?:ью|ьи|и|я)|сем(?:ью|ьи|и|я)\s+не\s+нуж|без\s+родств/.test(lower);
  const noWar = /без\s+во[йи]н|во[йи]н[уы]\s*н[еи]?\s*над|н[еи]?\s*над[оа]?\s*(?:про\s+)?во[йи]н|как\s+про\s+во[йи]н|не\s+пиши\s+во[йи]н|без\s+боев|без\s+воен|военн[а-я]*\s*н[еи]?\s*над/.test(lower);
  const noPolitics = /без\s+полит|политику\s+не\s+над|не\s+над[оа]?\s+полит/.test(lower);
  const onlySchool = /где\s+он\s+учил|где\s+она\s+учил|где\s+учил|где\s+училась|школ|учеб|образован|на\s+кого\s+пош|кем\s+стал|карьер|путь|деятельност/.test(lower);
  const raisedExplicit = /кто\s+воспит|воспит|родител|бабушк|дедушк|детств/.test(lower);
  const raisedSoft = /мам|пап|сем(?:ью|ьи|и|я)/.test(lower) && !noFamily;
  const onlyRaised = raisedExplicit || raisedSoft;
  const grade = (original.match(/(?:^|\s)(\d{1,2})\s*(?:класс|кл|класса)(?:\s|$)/i) || [])[1] || '';

  add(noFamily, 'Без раздела «Семья»');
  add(noWar, 'Без раздела войны и боевых действий');
  add(noPolitics, 'Без политики');
  add(onlySchool, 'Особое внимание: где учился и на кого пошёл');
  add(onlyRaised, 'Особое внимание: кто воспитывал и детство');

  return { original, lower, constraints, grade, noFamily, noWar, noPolitics, onlySchool, onlyRaised, mode };
}

function cleanNameQuery(raw, parsed) {
  let s = norm(raw);

  // Remove buttons/commands.
  s = s.replace(/^(\/start|📝\s*готовый\s*текст|📊\s*презентация)\s*/i, '');

  // Remove explicit constraint phrases before name extraction.
  s = s
    .replace(/без\s+сем[ьиюя][^\n]*/g, ' ')
    .replace(/семью\s+не\s+над[оа]?[^\n]*/g, ' ')
    .replace(/не\s+над[оа]?\s+сем[^\n]*/g, ' ')
    .replace(/без\s+во[йи]н[^\n]*/g, ' ')
    .replace(/войну\s+не\s+над[оа]?[^\n]*/g, ' ')
    .replace(/не\s+над[оа]?\s+во[йи]н[^\n]*/g, ' ')
    .replace(/как\s+про\s+во[йи]н[^\n]*/g, ' ')
    .replace(/без\s+полит[^\n]*/g, ' ')
    .replace(/где\s+он\s+учил[^\n]*/g, ' ')
    .replace(/где\s+она\s+учил[^\n]*/g, ' ')
    .replace(/где\s+учил[^\n]*/g, ' ')
    .replace(/на\s+кого\s+пош[ёе]л[^\n]*/g, ' ')
    .replace(/кто\s+воспитывал[^\n]*/g, ' ')
    .replace(/кто\s+его\s+воспитывал[^\n]*/g, ' ')
    .replace(/кто\s+её\s+воспитывал[^\n]*/g, ' ')
    .replace(/и\s+так\s+далее/g, ' ')
    .replace(/и\s+т\s*д/g, ' ')
    .replace(/так\s+далее/g, ' ')
    .replace(/\b\d{1,2}\s*(?:класс|кл|класса)\b/g, ' ')
    .replace(/\b(?:для|в)\s*\d{1,2}\s*(?:класс|кл|класса)\b/g, ' ');

  // Remove request/action words, keep only likely name words.
  const tokens = s.split(/\s+/).filter(Boolean);
  const kept = [];
  for (const t of tokens) {
    if (STOP_WORDS.has(t)) continue;
    if (/^\d+$/.test(t)) continue;
    const stem = stripWordEndings(t);
    if (STOP_WORDS.has(stem)) continue;
    if (t.length <= 2 && !/[а-яa-z]/.test(t)) continue;
    kept.push(t);
  }

  return kept.join(' ').trim();
}

function findLocal(query) {
  const q = norm(query);
  if (!q) return null;

  // Exact alias wins.
  if (ALIASES[q]) return LOCAL_PEOPLE.find(p => p.name === ALIASES[q]);

  // Also allow alias with ending (например, «пушкина»).
  for (const [alias, canonical] of Object.entries(ALIASES)) {
    const a = stripWordEndings(alias);
    const qq = stripWordEndings(q);
    if (a.length >= 4 && qq.length >= 4 && (a === qq || similarity(a, qq) > 0.9)) {
      return LOCAL_PEOPLE.find(p => p.name === canonical);
    }
  }

  const tokens = q.split(' ').filter(Boolean);
  let best = null;
  let bestScore = 0;
  for (const p of LOCAL_PEOPLE) {
    const candidateVariants = [p.name, ...p.aliases];
    let score = 0;
    for (const v of candidateVariants) {
      const vt = norm(v).split(' ').filter(Boolean);
      score = Math.max(score, bestNameScore(vt, tokens));
    }
    if (score > bestScore) { bestScore = score; best = p; }
  }

  // Require confident match; otherwise let online sources try.
  return bestScore >= 0.67 ? best : null;
}

// ============================ SOURCES ============================

const apiCache = new Map();
let apiQueue = Promise.resolve();
let lastCall = 0;

function apiGet(url, params, opts = {}) {
  const key = url + JSON.stringify(params || {});
  if (apiCache.has(key)) return apiCache.get(key);
  const run = apiQueue.then(async () => {
    const wait = Math.max(0, 380 - (Date.now() - lastCall));
    if (wait) await new Promise(r => setTimeout(r, wait));
    lastCall = Date.now();
    let lastErr = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const res = await http.get(url, { params: params || {}, timeout: opts.timeout || MAX_REQUEST_MS });
        apiCache.set(key, res.data);
        return res.data;
      } catch (e) {
        lastErr = e;
        const status = e && e.response && e.response.status;
        if ((status !== 429 && status < 500) || attempt === 1) break;
        const retryAfter = Number(e.response.headers['retry-after']) || 0;
        await new Promise(r => setTimeout(r, Math.min(2500, retryAfter ? retryAfter * 1000 : 700 * (attempt + 1))));
      }
    }
    throw lastErr;
  });
  apiQueue = run.catch(() => {});
  return run;
}

function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(`${label} timeout`)), ms))
  ]);
}

async function wikiSearch(base, name) {
  const data = await apiGet(base, {
    action: 'query', list: 'search', srsearch: name, srlimit: 6,
    format: 'json', origin: '*'
  });
  return (data.query && data.query.search || []).map(x => x.title);
}

async function wikiExtract(base, title) {
  const data = await apiGet(base, {
    action: 'query', prop: 'extracts|pageprops', exintro: 1, explaintext: 1,
    titles: title, redirects: 1, format: 'json', origin: '*'
  });
  const pages = data.query && data.query.pages || {};
  const page = Object.values(pages)[0] || {};
  return {
    title: page.title || title,
    text: String(page.extract || '').replace(/\s+\n/g, '\n').trim(),
    wikidata: page.pageprops && page.pageprops.wikibase_item || ''
  };
}

async function wikidataLabel(qid) {
  if (!qid) return '';
  try {
    const data = await apiGet(WIKIDATA_API, {
      action: 'wbgetentities', ids: qid, props: 'labels', languages: 'ru|en', format: 'json', origin: '*'
    });
    const ent = data.entities && data.entities[qid];
    return (ent && ent.labels && (ent.labels.ru || ent.labels.en) || {}).value || '';
  } catch (_) { return ''; }
}

async function ddgSearch(name) {
  try {
    const data = await apiGet('https://api.duckduckgo.com/', {
      q: name, format: 'json', no_html: 1, skip_disambig: 1
    });
    const text = [data.AbstractText, data.Answer].filter(Boolean).join('\n');
    return { title: data.Heading || name, text: text.trim() };
  } catch (_) { return null; }
}

function looksPerson(text, title) {
  const t = `${title || ''} ${String(text || '').slice(0, 800)}`;
  const bad = /КБЖД|железная дорога|станция|поезд|маршрут|остановка|автобус| Уикипедия /i;
  if (bad.test(t)) return false;
  const born = /родился|родилась|birth|born|\b18\d\d\b|\b19\d\d\b|\b20\d\d\b/i.test(t);
  const human = /человек|политик|писатель|поэт|учён|учен|физик|химик|биолог|президент|премьер|корол|царь|князь|генерал|художник|композитор|актёр|актер|певец|космонавт|философ|математик|инженер|изобретатель|деятел|персон/i.test(t);
  return born || human;
}

function deriveFacts(title, text) {
  const facts = { birth: '', death: '', born: '', country: '', job: '', education: '', career: '', raised: '', family: {} };
  const b = text.match(/(?:родился|родилась)[^.]{0,120}?(\d{4})/i);
  if (b) facts.birth = b[1];
  const bornPlace = text.match(/(?:родился|родилась)[^.]{0,120}?\bв\s+([^,.]{2,60})/i);
  if (bornPlace) facts.born = bornPlace[1].trim();
  const country = text.match(/(гражданство|государство|страна)[^:]{0,40}:\s*([^.\n]{2,50})/i);
  if (country) facts.country = country[2].trim();
  const job = text.match(/[^.]{0,80}(?:является|был|была| работает|работал|работала)\s+([^,.]{3,60})/i);
  if (job) facts.job = job[1].trim();
  const edu = text.match(/(?:учился|училась|окончил|окончила|образование)[^.]{0,160}/i);
  if (edu) facts.education = edu[0].trim();
  const career = text.match(/(?:карьера|деятельность|работал|работала|занимался|занималась)[^.]{0,180}/i);
  if (career) facts.career = career[0].trim();
  const father = text.match(/отец[^\n.]{0,80}/i);
  const mother = text.match(/мать[^\n.]{0,80}/i);
  if (father) facts.family['Отец'] = [father[0].replace(/^отец\s*/i, '').trim()];
  if (mother) facts.family['Мать'] = [mother[0].replace(/^мать\s*/i, '').trim()];
  return facts;
}

async function collectPerson(rawQuery, parsed) {
  const query = cleanNameQuery(rawQuery, parsed);
  if (!query) return null;

  const local = findLocal(query);
  if (local) return { ...local, title: local.name, source: 'проверенная локальная база' };

  const deadline = Date.now() + MAX_SEARCH_MS;
  const msLeft = () => Math.max(1000, deadline - Date.now());

  const bases = [WIKI_API, WIKI_FALLBACK];
  let wikiHit = null;

  for (const base of bases) {
    if (Date.now() > deadline) break;
    try {
      const titles = await withTimeout(wikiSearch(base, query), Math.min(10000, msLeft()), 'wiki search');
      for (const title of titles.slice(0, 4)) {
        if (Date.now() > deadline) break;
        if (/список|категория|шаблон|портал/i.test(title)) continue;
        const page = await withTimeout(wikiExtract(base, title), Math.min(10000, msLeft()), 'wiki extract');
        if (page.text && looksPerson(page.text, page.title)) {
          wikiHit = { ...page, source: 'Википедия' };
          break;
        }
      }
    } catch (_) {}
    if (wikiHit) break;
  }

  let wdHit = null;
  if (wikiHit && wikiHit.wikidata && Date.now() < deadline) {
    const label = await withTimeout(wikidataLabel(wikiHit.wikidata), Math.min(6000, msLeft()), 'wikidata').catch(() => '');
    if (label && looksPerson(label, label)) wdHit = label;
  }

  let ddgHit = null;
  if ((!wikiHit || !looksPerson(wikiHit.text, wikiHit.title)) && Date.now() < deadline) {
    ddgHit = await withTimeout(ddgSearch(query), Math.min(8000, msLeft()), 'ddg').catch(() => null);
  }

  const candidates = [];
  if (wikiHit) candidates.push({ title: wikiHit.title, text: wikiHit.text, source: 'Википедия', score: 3 });
  if (wdHit) candidates.push({ title: wdHit, text: '', source: 'Wikidata', score: 1 });
  if (ddgHit && ddgHit.text && looksPerson(ddgHit.text, ddgHit.title)) {
    candidates.push({ title: ddgHit.title || query, text: ddgHit.text, source: 'DuckDuckGo', score: 2 });
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) => b.score - a.score);
  const best = candidates[0];
  return {
    name: best.title,
    title: best.title,
    text: best.text || `${best.title} — общественная личность. Подробная биография доступна в открытых источниках.`,
    facts: deriveFacts(best.title, best.text),
    source: candidates.map(c => c.source).join(' + ')
  };
}

// ============================ REPORT ============================

function makeReport(d, note, parsed) {
  const f = d.facts || {};
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. ${f.birth ? 'Родился(ась) в ' + f.birth + ' году' : 'Информация о годе рождения приведена в открытых источниках'}${f.born ? ' в ' + f.born : ''}. ${f.country ? 'Гражданство: ' + f.country + '. ' : ''}${f.job ? 'Род деятельности: ' + dot(f.job) : ''}`);
  lines.push('', 'Основная часть');
  let ss = sentences(d.text || '', 8);
  if (parsed && parsed.noWar) ss = ss.filter(s => !/войн|боев|военн/i.test(s));
  if (parsed && parsed.noPolitics) ss = ss.filter(s => !/полит|парти|президент|правительств/i.test(s));
  if (ss.length) for (const s of ss) lines.push(s);
  else lines.push('В открытых источниках приведены основные сведения о жизни и деятельности этой личности.');

  if (f.education || f.career || (parsed && parsed.onlySchool)) {
    lines.push('', 'Образование и путь');
    if (f.education) lines.push('Образование: ' + dot(String(f.education).replace(/^.*?(учился|училась|окончил|окончила|образование)\s*/i, '')));
    if (f.career) lines.push('Деятельность: ' + dot(String(f.career).replace(/^.*?(деятельность|работал|работала|занимался|занималась)\s*/i, '')));
    if (!f.education && !f.career) lines.push('Подробные сведения об образовании и карьере приведены в открытых источниках.');
  }

  if (parsed && parsed.onlyRaised) {
    lines.push('', 'Воспитание');
    lines.push(f.raised ? dot(cap(f.raised)) : 'Сведения о воспитании приведены в открытых источниках.');
    const parents = Object.entries(f.family || {}).filter(([k]) => /отец|мать|родител/i.test(k));
    for (const [k, v] of parents) lines.push(`${k}: ${v.join(', ')}.`);
  }

  const fam = Object.entries(f.family || {});
  if (fam.length && !(parsed && parsed.noFamily)) {
    lines.push('', 'Семья');
    for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}.`);
  }

  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);

  let text = lines.join('\n');
  text = text.replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\s+([.,;:])/g, '$1')
    .trim();
  return text;
}

function sentences(text, limit = 6) {
  return String(text || '').replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ============================ PRESENTATION ============================

const THEMES = [
  { name: 'синий', bg: 'F2F6FF', dark: '1E3A8A', accent: '3B82F6', light: 'DBEAFE' },
  { name: 'зелёный', bg: 'F1F8F2', dark: '1B5E20', accent: '43A047', light: 'DCEDC8' },
  { name: 'фиолетовый', bg: 'F7F3FF', dark: '4A148C', accent: '7B1FA2', light: 'E1BEE7' },
  { name: 'granat', bg: 'F4F6FA', dark: '0F2A5A', accent: '2E5CA8', light: 'D6E4FF' },
  { name: 'тёплый', bg: 'FFF8EE', dark: '8A3B12', accent: 'D97706', light: 'FFE0B2' }
];

function pickTheme(title) {
  const h = [...norm(title)].reduce((a, c) => a + c.charCodeAt(0), 0);
  return THEMES[h % THEMES.length];
}

function splitBullets(text, max = 5) {
  const ss = sentences(text, max);
  const out = [];
  for (const s of ss) {
    let clean = s.replace(/^.*?(?:родился|родилась)[^,]{0,60},\s*/i, '')
      .replace(/^В\s+\d{4}\s+году,\s*/i, '');
    if (clean.length > 180) clean = clean.slice(0, 177).trim() + '…';
    out.push(clean || s);
  }
  return out.slice(0, max);
}

function buildSlidePlan(d, slideCount, parsed) {
  const n = Math.max(3, Math.min(16, Number(slideCount) || 7));
  const theme = pickTheme(d.title);
  const f = d.facts || {};

  const plan = [];
  plan.push({
    kind: 'title',
    title: d.title,
    subtitle: 'Доклад / презентация',
    badge: `${parsed.grade ? parsed.grade + ' класс' : 'Школьный проект'}`
  });

  plan.push({
    kind: 'bullets',
    title: 'Кто это',
    bullets: [
      `${d.title} — ${f.job || 'известная личность'}.`,
      f.birth ? `Год рождения: ${f.birth}.` : 'Сведения о жизни приведены в открытых источниках.',
      f.born ? `Место рождения: ${f.born}.` : '',
      f.country ? `Страна: ${f.country}.` : ''
    ].filter(Boolean)
  });

  const bio = splitBullets(d.text || '', 5);
  const bioChunks = [];
  for (let i = 0; i < bio.length; i += 3) bioChunks.push(bio.slice(i, i + 3));
  if (!bioChunks.length) bioChunks.push(['Основные сведения приведены в открытых источниках.']);

  for (const chunk of bioChunks) {
    plan.push({ kind: 'bullets', title: 'Биография', bullets: chunk });
  }

  if (f.education || f.career || parsed.onlySchool) {
    plan.push({
      kind: 'bullets',
      title: 'Образование и путь',
      bullets: [
        f.education ? 'Где учился(ась): ' + String(f.education).replace(/^.*?(учился|училась|окончил|окончила|образование)\s*/i, '') : 'Сведения об образовании приведены в открытых источниках.',
        f.career ? 'На кого пошёл(шла): ' + String(f.career).replace(/^.*?(деятельность|работал|работала|занимался|занималась)\s*/i, '') : 'После учёбы связал(а) жизнь с общественной деятельностью.'
      ].filter(Boolean)
    });
  }

  if (f.raised || parsed.onlyRaised) {
    plan.push({
      kind: 'bullets',
      title: 'Детство и воспитание',
      bullets: [
        f.raised || 'Воспитывался(ась) в семье, которая повлияла на его/её взгляды и интересы.',
        f.family && f.family['Отец'] ? 'Отец: ' + f.family['Отец'].join(', ') : '',
        f.family && f.family['Мать'] ? 'Мать: ' + f.family['Мать'].join(', ') : ''
      ].filter(Boolean)
    });
  }

  const fam = Object.entries(f.family || {});
  if (fam.length && !parsed.noFamily) {
    plan.push({
      kind: 'bullets',
      title: 'Семья',
      bullets: fam.slice(0, 4).map(([k, v]) => `${k}: ${v.slice(0, 2).join(', ')}`)
    });
  }

  plan.push({
    kind: 'bullets',
    title: 'Значение',
    bullets: [
      `${d.title} оставил(а) заметный след в истории.`,
      'Его/её жизнь и идеи изучают в школе.',
      'Биография подходит для доклада и самостоятельной работы.'
    ].filter(Boolean)
  });

  plan.push({
    kind: 'final',
    title: 'Вывод',
    bullets: [
      `${d.title} — личность, чья жизнь связана с важными событиями и идеями.`,
      parsed.noWar ? 'В докладе не рассматривались военные действия.' : '',
      parsed.noFamily ? 'Семья не рассматривалась по заданию.' : ''
    ].filter(Boolean)
  });

  // Fill or trim to requested count while preserving logical order.
  while (plan.length < n) {
    const idx = plan.length;
    const chunk = bio[(idx - 3) % Math.max(1, bio.length)] || 'Дополнительные сведения можно найти в открытых источниках.';
    plan.splice(Math.max(3, plan.length - 2), 0, { kind: 'bullets', title: 'Подробности', bullets: [chunk] });
  }
  if (plan.length > n) {
    const fixed = plan.filter(p => p.kind !== 'bullets').length;
    const keep = Math.max(0, n - fixed);
    const bullets = plan.filter(p => p.kind === 'bullets');
    const other = plan.filter(p => p.kind !== 'bullets');
    const rebuilt = [];
    let bi = 0;
    for (const p of plan) {
      if (p.kind === 'bullets') {
        if (bi++ < keep) rebuilt.push(p);
      } else rebuilt.push(p);
    }
    return rebuilt.slice(0, n);
  }
  return plan;
}

async function makePptx(d, slideCount, parsed) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'SchoolBot';
  pptx.subject = 'Школьный доклад';
  pptx.title = d.title;
  pptx.lang = 'ru-RU';
  pptx.theme = {
    headFontFace: 'Arial', bodyFontFace: 'Arial', lang: 'ru-RU'
  };

  const plan = buildSlidePlan(d, slideCount, parsed);
  const theme = pickTheme(d.title);

  for (const item of plan) {
    const slide = pptx.addSlide();
    slide.background = { color: theme.bg };

    // Top colored band (slide size is 10 x 5.625 in).
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: '100%', h: 0.12, fill: { color: theme.accent }, line: { color: theme.accent }
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0.12, w: '100%', h: 0.05, fill: { color: theme.light }, line: { color: theme.light }
    });

    if (item.kind === 'title') {
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.5, y: 0.5, w: '90%', h: 1.0,
        fill: { color: theme.dark }, line: { color: theme.dark }, radius: 0.12
      });
      slide.addText(item.title, {
        x: 0.65, y: 0.63, w: '85%', h: 0.75, align: 'center', valign: 'mid',
        fontSize: 30, bold: true, color: 'FFFFFF', fontFace: 'Arial', margin: 0
      });
      if (item.subtitle) slide.addText(item.subtitle, {
        x: 0, y: 4.35, w: '100%', h: 0.35, align: 'center', fontSize: 16, color: theme.dark
      });
      slide.addText(item.badge || '', {
        x: 0, y: 4.85, w: '100%', h: 0.35, align: 'center', fontSize: 14,
        color: theme.dark, fontFace: 'Arial'
      });
      continue;
    }

    // Regular slide: centered top heading.
    slide.addText(item.title, {
      x: 0.5, y: 0.3, w: '90%', h: 0.6,
      align: 'center', valign: 'mid', fontSize: 24, bold: true,
      color: theme.dark, fontFace: 'Arial', margin: 0
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 2.0, y: 0.95, w: 6.0, h: 0.04, fill: { color: theme.accent }, line: { color: theme.accent }
    });

    const bullets = (item.bullets || []).filter(Boolean);
    const fontSize = bullets.length <= 2 ? 24 : bullets.length <= 4 ? 21 : 18;
    const bodyY = 1.15;
    const bodyH = bullets.length <= 2 ? 3.5 : 4.15;

    slide.addText(
      bullets.map((b, i) => ({ text: String(b).replace(/\s+/g, ' ').trim(), options: { bullet: { code: '2022' }, breakLine: i < bullets.length - 1 } })),
      {
        x: 0.7, y: bodyY, w: '86%', h: bodyH, valign: 'mid', align: 'center',
        fontSize, color: '1F2937', fontFace: 'Arial', lineSpacingMultiple: 1.18,
        paraSpaceAfterPt: 12, margin: 0.08, fit: 'shrink'
      }
    );

    // Decorative accent element for a modern, non-plain look.
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 9.2, y: 5.05, w: 0.25, h: 0.25, fill: { color: theme.accent, transparency: 40 }, line: { color: theme.accent }
    });
  }

  const outDir = path.join(__dirname, 'generated');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const file = path.join(outDir, `presentation_${Date.now()}.pptx`);
  await pptx.writeFile({ fileName: file });
  return file;
}

// ============================ TELEGRAM ============================

const bot = new Telegraf(TOKEN);

const mainMenu = Markup.keyboard([
  ['📝 Готовый текст', '📊 Презентация']
]).resize().reply_markup;

const removeKb = Markup.removeKeyboard().reply_markup;

const sessions = new Map();
function getSession(id) {
  if (!sessions.has(id)) sessions.set(id, { step: 'menu', awaitingPres: false, name: '', note: '' });
  return sessions.get(id);
}

async function safeReply(ctx, text, markup) {
  const chunks = String(text).match(/[\s\S]{1,3800}/g) || [text];
  for (const chunk of chunks) {
    try { await ctx.reply(chunk, markup ? { reply_markup: markup } : undefined); }
    catch (e) { console.error('reply error', e.message); }
  }
}

async function handleTextRequest(ctx, s, raw) {
  const parsed = parseRequest(raw, 'text');
  await safeReply(ctx, 'Понял запрос. Ищу и сверяю по нескольким источникам (до 30 секунд)…');

  let person = null;
  try {
    person = await withTimeout(collectPerson(raw, parsed), MAX_SEARCH_MS, 'search');
  } catch (e) {
    console.error('search error', e.message);
  }

  if (!person) {
    await safeReply(ctx, 'Не смог понять, о ком речь. Напиши, пожалуйста, имя человека целиком, например: «Альберт Эйнштейн» или «Уинстон Черчилль».');
    return;
  }

  const report = makeReport(person, raw, parsed);
  s.name = person.title;
  s.note = raw;
  s.step = 'menu';
  await safeReply(ctx, report);
  await safeReply(ctx, `Источник: ${person.source}. Если нужно изменить доклад, напиши уточнение (например, «без войны», «8 класс», «где учился»).`, mainMenu);
}

async function handlePresRequest(ctx, s, raw) {
  const parsed = parseRequest(raw, 'pres');
  s.storedParsed = parsed;

  const nums = String(raw).match(/\b(\d{1,2})\s*(?:слайд|слайдов|слайда|лист|листов)\b/i) ||
    String(raw).match(/\b(\d{1,2})\s*[.)]?\s*$/);
  s.slideCount = nums ? Math.max(3, Math.min(16, Number(nums[1]))) : 0;

  if (!s.slideCount) {
    s.step = 'pres_slides';
    await safeReply(ctx, 'Сколько слайдов нужно? Напиши числом от 3 до 16 (например, 8).', removeKb);
    return;
  }

  await generatePres(ctx, s, raw);
}

async function generatePres(ctx, s, raw) {
  const parsed = s.storedParsed || parseRequest(raw || s.name || '', 'pres');
  await safeReply(ctx, 'Готовлю красивую презентацию…');

  let person = null;
  try {
    const query = s.name || raw || '';
    person = await withTimeout(collectPerson(query, parsed), MAX_SEARCH_MS, 'search');
  } catch (_) {}

  if (!person) {
    await safeReply(ctx, 'Не смог понять, о ком презентация. Напиши имя целиком.', mainMenu);
    s.step = 'menu';
    return;
  }

  try {
    const file = await makePptx(person, s.slideCount, parsed);
    await ctx.replyWithDocument({ source: file, filename: `Презентация_${person.title.replace(/[^\wа-яА-ЯёЁ-]+/g, '_')}.pptx` });
    await safeReply(ctx, `Готово! Слайдов: ${s.slideCount}. Оформление подобрано автоматически.`, mainMenu);
  } catch (e) {
    console.error('pptx error', e);
    await safeReply(ctx, 'Не удалось собрать файл презентации. Попробуй ещё раз.', mainMenu);
  }
  s.step = 'menu';
}

bot.on('text', async (ctx, next) => {
  if (!ctx.message || !ctx.message.text) return next();
  const text = ctx.message.text.trim();
  const s = getSession(ctx.from.id);

  if (text === '/start') {
    s.step = 'menu';
    await safeReply(ctx, `Привет! Я помогаю делать школьные доклады и презентации о публичных личностях. Версия ${BOT_VERSION}.\nВыбери кнопку или просто напиши, про кого нужно.`, mainMenu);
    return;
  }

  if (text === '📝 Готовый текст') {
    s.step = 'text_topic';
    await safeReply(ctx, 'Напиши, про кого и что учесть. Например:\nУинстон Черчилль, 8 класс, без войны, где учился и на кого пошёл', removeKb);
    return;
  }

  if (text === '📊 Презентация') {
    s.step = 'pres_topic';
    await safeReply(ctx, 'Напиши, про кого презентация и сколько слайдов. Например:\nУинстон Черчилль, 8 слайдов', removeKb);
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    s.name = text;
    await handlePresRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_slides') {
    await generatePres(ctx, s, text);
    return;
  }

  // Free-form message from menu: if it looks like a presentation request, make pptx; else report.
  if (/презентац|слайд|презент/i.test(text)) {
    s.name = text;
    await handlePresRequest(ctx, s, text);
  } else {
    await handleTextRequest(ctx, s, text);
  }
});

bot.on('document', ctx => ctx.reply('Пришли, пожалуйста, текстом имя человека, про которого нужен доклад.'));

// ============================ START ============================

async function start() {
  if (WEBHOOK_URL) {
    const app = express();
    app.use(express.json());
    app.post('/webhook', (req, res) => {
      bot.handleUpdate(req.body).catch(e => console.error('webhook error', e));
      res.sendStatus(200);
    });
    app.get('/', (_, res) => res.send(`Bot is alive. ${BOT_VERSION}`));
    app.listen(PORT, () => console.log('Webhook server on port', PORT));
    await bot.telegram.setWebhook(`${WEBHOOK_URL}/webhook`);
    console.log('Webhook set');
  } else {
    await bot.telegram.deleteWebhook({ drop_pending_updates: true });
    await bot.launch({ dropPendingUpdates: true });
    console.log('Polling started');
  }
}

if (require.main === module) {
  start().catch(e => {
    console.error('Start failed', e);
    process.exit(1);
  });

  process.once('SIGINT', () => bot.stop('SIGINT'));
  process.once('SIGTERM', () => bot.stop('SIGTERM'));
}

module.exports = {
  parseRequest, cleanNameQuery, findLocal, collectPerson,
  makeReport, makePptx, buildSlidePlan, LOCAL_PEOPLE
};


// --- Защита от лимитов Википедии/Викиданных (ошибка 429) ---
// Render использует общий IP, поэтому запросы идут реже, повторяются при 429
// и кэшируются, чтобы повторный вопрос про того же человека не тратил лимит.
const apiCache = new Map();
let lastApiCall = 0;
let apiBusy = false;

async function apiGet(url, params) {
  const key = url + '|' + JSON.stringify(params || {});
  if (apiCache.has(key)) return apiCache.get(key);

  // Несколько вызовов подряд (поиск в Википедии + Викиданных) идут строго по очереди.
  while (apiBusy) await new Promise(r => setTimeout(r, 120));
  apiBusy = true;
  try {
    const gap = 400 - (Date.now() - lastApiCall);
    if (gap > 0) await new Promise(r => setTimeout(r, gap));

    let lastErr;
    for (let attempt = 0; attempt < 5; attempt++) {
      lastApiCall = Date.now();
      try {
        const { data } = await axios.get(url, { params, timeout: 30000, ...netOpts });
        apiCache.set(key, data);
        return data;
      } catch (e) {
        lastErr = e;
        const status = e.response?.status;
        const retryable = status === 429 || status === 500 || status === 502 ||
                          status === 503 || status === 504 || !status;
        if (!retryable || attempt === 4) break;
        const header = Number(e.response?.headers?.['retry-after']);
        const wait = Number.isFinite(header) && header > 0
          ? header * 1000
          : Math.min(2000 * Math.pow(2, attempt), 15000);
        await new Promise(r => setTimeout(r, wait));
      }
    }
    throw lastErr;
  } finally {
    apiBusy = false;
  }
}

const TEXT_BTN = '📝 Готовый текст';
const PRES_BTN = '📊 Презентация';

const mainMenu = Markup.keyboard([[TEXT_BTN, PRES_BTN]]).resize();
const removeKb = Markup.removeKeyboard();

const YES_WORDS = ['да', 'давай', 'давай же', 'ага', 'угу', 'конечно', 'ок', 'окей',
  'ладно', 'хочу', 'сделай', 'презентацию', 'презентация', 'yes', 'yeah', 'добре'];
const NO_WORDS = ['нет', 'не надо', 'не хочу', 'не сейчас', 'позже', 'стоп', 'отмена', 'no'];

const yesNoKeyboard = () => Markup.inlineKeyboard([
  [Markup.button.callback('Да, сделать презентацию', 'pres:yes'),
   Markup.button.callback('Нет, спасибо', 'pres:no')],
]);

const STYLES = { classic: 'Классический', modern: 'Современный', kids: 'Детский' };
const COLORS = {
  blue: ['Синий', '1F4E8C'],
  green: ['Зелёный', '226E4E'],
  red: ['Красный', '962D37'],
  violet: ['Фиолетовый', '5E378C'],
};

const styleKeyboard = () => Markup.inlineKeyboard(
  Object.entries(STYLES).map(([k, label]) => Markup.button.callback(label, `style:${k}`))
);
const colorKeyboard = () => Markup.inlineKeyboard(
  Object.entries(COLORS).map(([k, v]) => Markup.button.callback(v[0], `color:${k}`))
);

// Простейшие сессии в памяти
const sessions = new Map();
function session(ctx) {
  const id = ctx.chat.id;
  if (!sessions.has(id)) sessions.set(id, { step: 'menu' });
  return sessions.get(id);
}

// ---------- Открытые данные ----------

// --- Понимание разговорных запросов ("напиши про энштейна" и т.п.) ---
const STOP_WORDS = new Set([
  'напиши','написать','напишите','про','об','about','сделай','сделать','сделайте',
  'давай','дай','пришли','инфа','инфу','информация','информацию',
  'текст','доклад','доклада','презентация','презентацию','презентации',
  'биография','биографию','биографии','расскажи','рассказать','кто','такой','такая',
  'что','какой','какая','нужно','надо','помоги','найди','ищи',
  'поищи','сведения','данные','материал','материалы','тему','тема','для','класс',
  'класса','школе','школа','школь','человек','личность','жизнь','жизни'
]);

// Частые сокращения и опечатки -> полное имя.
const ALIASES = {
  'энштейн': 'альберт эйнштейн', 'энштейна': 'альберт эйнштейн',
  'эйнштейн': 'альберт эйнштейн', 'эйнштейна': 'альберт эйнштейн', 'айнштейн': 'альберт эйнштейн',
  'пушкин': 'александр пушкин', 'пушкина': 'александр пушкин',
  'лермонтов': 'михаил лермонтов', 'толстой': 'лев толстой', 'чехов': 'антон чехов',
  'достоевский': 'фёдор достоевский', 'гоголь': 'николай гоголь',
  'черчилль': 'уинстон черчилль', 'наполеон': 'наполеон бонапарт',
  'сталин': 'иосиф сталин', 'ленин': 'владимир ленин', 'гитлер': 'адольф гитлер',
  'петр 1': 'пётр великий', 'петр великий': 'пётр великий', 'пётр 1': 'пётр великий',
  'екатерина 2': 'екатерина великая', 'екатерина великая': 'екатерина великая',
  'гагарин': 'юрий гагарин', 'королев': 'сергей королёв', 'менделеев': 'дмитрий менделеев',
  'ломоносов': 'михаил ломоносов', 'суворов': 'александр суворов', 'жуков': 'георгий жуков'
};

// Встроенная база известных личностей. Работает даже когда Википедия недоступна
// (у Render общий IP, и Википедия иногда отвечает 429).
const LOCAL_PEOPLE = {
  'альберт эйнштейн': {
    title: 'Альберт Эйнштейн',
    text: 'Альберт Эйнштейн (1879—1955) — выдающийся физик-теоретик, создатель специальной и общей теории относительности. Родился в Ульме в Германии, работал в Швейцарии, Германии и США. В 1905 году опубликовал работы по фотоэффекту, броуновскому движению и специальной теории относительности. В 1915 году завершил общую теорию относительности, а в 1921 году получил Нобелевскую премию по физике. Его имя стало символом современной науки.',
    facts: { birth: '1879', death: '1955', born: 'Ульм', died: 'Принстон', country: 'Германия, США', job: 'физик' },
    family: { 'Отец': ['Герман Эйнштейн'], 'Мать': ['Паулина Кох'], 'Супруг(а)': ['Милева Марич', 'Эльза Лёвенталь'] }
  },
  'александр пушкин': {
    title: 'Александр Пушкин',
    text: 'Александр Сергеевич Пушкин (1799—1837) — великий русский поэт и писатель, основоположник современной русской литературы. Родился в Москве в дворянской семье. Учился в Царскосельском лицее, где начал писать стихи. Ему принадлежат «Евгений Онегин», «Борис Годунов», «Капитанская дочка», «Руслан и Людмила» и многие другие стихотворения. Творчество Пушкина оказало огромное влияние на русскую культуру. Погиб на дуэли с Дантесом в 37 лет.',
    facts: { birth: '1799', death: '1837', born: 'Москва', died: 'Санкт-Петербург', country: 'Россия', job: 'поэт, писатель' },
    family: { 'Отец': ['Сергей Львович Пушкин'], 'Мать': ['Надежда Осиповна Пушкина'], 'Супруг(а)': ['Наталья Гончарова'] }
  },
  'михаил лермонтов': {
    title: 'Михаил Лермонтов',
    text: 'Михаил Юрьевич Лермонтов (1814—1841) — русский поэт и писатель романтического направления. Родился в Москве, детство провёл в имении бабушки в Пятигорске. В 1837 году написал стихотворение «На смерть поэта», посвящённое гибели Пушкина, и был сослан на Кавказ. Его главные произведения — роман «Герой нашего времени», поэмы «Демон» и «Мцыри», стихотворение «Парус». Погиб на дуэли в Пятигорске в возрасте 26 лет.',
    facts: { birth: '1814', death: '1841', born: 'Москва', died: 'Пятигорск', country: 'Россия', job: 'поэт, писатель' },
    family: { 'Отец': ['Юрий Петрович Лермонтов'], 'Мать': ['Мария Михайловна Арсеньева'] }
  },
  'лев толстой': {
    title: 'Лев Толстой',
    text: 'Лев Николаевич Толстой (1828—1910) — великий русский писатель и мыслитель. Родился в имении Ясная Поляна под Тулой. Участвовал в Крымской войне, затем целиком посвятил себя литературе. Ему принадлежат романы «Война и мир», «Анна Каренина», «Воскресение», а также рассказы и философские работы. Толстой проповедовал непротивление злу насилием и простую жизнь. Умер на станции Астапово, уйдя из дома в поисках духовной правды.',
    facts: { birth: '1828', death: '1910', born: 'Ясная Поляна', died: 'Астапово', country: 'Россия', job: 'писатель' },
    family: { 'Отец': ['Николай Ильич Толстой'], 'Мать': ['Мария Дмитриевна Толстая'], 'Супруг(а)': ['София Толстая'] }
  },
  'антон чехов': {
    title: 'Антон Чехов',
    text: 'Антон Павлович Чехов (1860—1904) — русский писатель, драматург и врач. Родился в Таганроге в семье лавочника. Окончил медицинский факультет Московского университета и практиковал как врач, что повлияло на его произведения. написал рассказы «Хамелеон», «Толстый и тонкий», «Человек в футляре», пьесы «Чайка», «Дядя Ваня», «Три сестры», «Вишнёвый сад». Его герои — обычные люди, а стиль отличается краткостью и тонким юмором. Умер от туберкулёза в Германии.',
    facts: { birth: '1860', death: '1904', born: 'Таганрог', died: 'Бадёнвайлер', country: 'Россия', job: 'писатель, врач' },
    family: { 'Отец': ['Павел Егорович Чехов'], 'Мать': ['Евгения Яковлевна Чехова'], 'Супруг(а)': ['Ольга Книппер-Чехова'] }
  },
  'фёдор достоевский': {
    title: 'Фёдор Достоевский',
    text: 'Фёдор Михайлович Достоевский (1821—1881) — русский писатель, один из величайших романистов мира. Родился в Москве в семье врача. В 1849 году был арестован за участие в кружке петрашевцев и приговорён к каторге в Сибири, что стало переломом в его жизни. Главные романы — «Преступление и наказание», «Идиот», «Братья Карамазовы», «Бесы». Его произведения исследуют вопросы совести, веры и человеческой свободы.',
    facts: { birth: '1821', death: '1881', born: 'Москва', died: 'Санкт-Петербург', country: 'Россия', job: 'писатель' },
    family: { 'Отец': ['Михаил Андреевич Достоевский'], 'Супруг(а)': ['Анна Достоевская'] }
  },
  'николай гоголь': {
    title: 'Николай Гоголь',
    text: 'Николай Васильевич Гоголь (1809—1852) — русский писатель и драматург, классик мировой литературы. Родился в местечке Сорочинцы на Украине. Переехав в Петербург, стал известен после «Вечеров на хуторе близ Диканьки». Ему принадлежат повести «Тарас Бульба», «Шинель», «Мёртвые души», пьеса «Ревизор». Гоголь высмеивал пороки общества и создал особый мир «маленьких людей». Умер в Москве, работая над второй частью «Мёртвых душ».',
    facts: { birth: '1809', death: '1852', born: 'Сорочинцы', died: 'Москва', country: 'Россия', job: 'писатель, драматург' },
    family: { 'Отец': ['Василий Афанасьевич Гоголь-Яновский'], 'Мать': ['Мария Ивановна Гоголь'] }
  },
  'уинстон черчилль': {
    title: 'Уинстон Черчилль',
    text: 'Уинстон Черчилль (1874—1965) — британский государственный деятель, премьер-министр Великобритании в годы Второй мировой войны. Родился в Оксфордшире в аристократической семье. Был офицером, журналистом и политиком. В 1940 году возглавил правительство и выступил с знаменитой речью «Мы будем сражаться на берегах». После войны стал лауреатом Нобелевской премии по литературе. Считается одним из главных символов борьбы с фашизмом.',
    facts: { birth: '1874', death: '1965', born: 'Бленхейм', died: 'Лондон', country: 'Великобритания', job: 'политик, писатель' },
    family: { 'Отец': ['Рэндольф Черчилль'], 'Мать': ['Дженни Джером'], 'Супруг(а)': ['Клементина Черчилль'] }
  },
  'наполеон бонапарт': {
    title: 'Наполеон Бонапарт',
    text: 'Наполеон Бонапарт (1769—1821) — французский полководец и император. Родился на острове Корсика в дворянской семье итальянского происхождения. Выдвинулся во время Французской революции, командовал армией в Италии и Египте. В 1804 году провозгласил себя императором французов. Провёл ряд крупных реформ, в том числе принял Гражданский кодекс. После поражения при Ватерлоо в 1815 году был сослан на остров Святой Елены, где и умер.',
    facts: { birth: '1769', death: '1821', born: 'Аяччо', died: 'Святая Елена', country: 'Франция', job: 'полководец, император' },
    family: { 'Отец': ['Карло Бонапарт'], 'Мать': ['Летиция Бонапарт'], 'Супруг(а)': ['Жозефина Богарне', 'Мария Луиза Австрийская'] }
  },
  'иосиф сталин': {
    title: 'Иосиф Сталин',
    text: 'Иосиф Виссарионович Сталин (Джугашвили, 1878—1953) — советский политический деятель, руководитель СССР в 1920-е — 1950-е годы. Родился в Гори в Грузии в семье сапожника. Участвовал в революционном движении, был соратником Ленина. После его смерти возглавил страну, провёл индустриализацию и коллективизацию, а также сосредоточил огромную власть в годы репрессий. В годы Великой Отечественной войны был Верховным главнокомандующим. Его правление остаётся спорной и трагической страницей истории.',
    facts: { birth: '1878', death: '1953', born: 'Гори', died: 'Москва', country: 'СССР', job: 'политик' },
    family: { 'Отец': ['Виссарион Джугашвили'], 'Мать': ['Екатерина Геладзе'], 'Супруг(а)': ['Екатерина Сванидзе', 'Надежда Аллилуева'] }
  },
  'владимир ленин': {
    title: 'Владимир Ленин',
    text: 'Владимир Ильич Ленин (Ульянов, 1870—1924) — революционер, основатель Советского государства. Родился в Симбирске в семье учителя. Изучал право, увлёкся марксизмом и создал партию большевиков. В 1917 году возглавил Октябрьскую революцию и стал первым руководителем советской власти. Провёл национализацию банков и земли, ввёл новую экономическую политику. Его идеи оказали огромное влияние на мировую историю XX века.',
    facts: { birth: '1870', death: '1924', born: 'Симбирск', died: 'Горки', country: 'Россия, СССР', job: 'политик, революционер' },
    family: { 'Отец': ['Илья Николаевич Ульянов'], 'Мать': ['Мария Александровна Ульянова'], 'Супруг(а)': ['Надежда Крупская'] }
  },
  'пётр великий': {
    title: 'Пётр I Великий',
    text: 'Пётр I (1672—1725) — русский царь, а с 1721 года император, реформатор и полководец. Сын царя Алексея Михайловича. С детства увлекался военным делом и мореплаванием, ездил учиться в Голландию и Англию. Провёл масштабные реформы армии, флота и управления, основал Санкт-Петербург. Выиграл Северную войну со Швецией, получив выход к Балтийскому морю. Его правление сделало Россию морской державой, но реформы проводились жёстко.',
    facts: { birth: '1672', death: '1725', born: 'Москва', died: 'Санкт-Петербург', country: 'Россия', job: 'царь, император' },
    family: { 'Отец': ['Алексей Михайлович'], 'Мать': ['Наталья Кирилловна Нарышкина'], 'Супруг(а)': ['Евдокия Лопухина', 'Екатерина I'] }
  },
  'екатерина великая': {
    title: 'Екатерина II Великая',
    text: 'Екатерина II (1729—1796) — российская императрица, правившая с 1762 года. Урождённая принцесса София Фридерика Ангальт-Цербстская, приехала в Россию и вышла замуж за Петра III. После дворцового переворота стала императрицей. Проводила реформы управления и образования, покровительствовала искусствам, переписывалась с Вольтером. При ней Россия расширила территорию за счёт Крыма и Причерноморья. Время её правления называют «золотым веком» русского дворянства.',
    facts: { birth: '1729', death: '1796', born: 'Штеттин', died: 'Санкт-Петербург', country: 'Россия', job: 'императрица' },
    family: { 'Супруг(а)': ['Пётр III'] }
  },
  'юрий гагарин': {
    title: 'Юрий Гагарин',
    text: 'Юрий Алексеевич Гагарин (1934—1968) — первый космонавт планеты. Родился в деревне Клушино под Смоленском в колхозной семье. Учился в техническом училище и авиационном институте, служил лётчиком. 12 апреля 1961 года на корабле «Восток-1» совершил первый в истории полёт в космос, облетев Землю за 108 минут. Его слова «Поехали!» стали крылатыми. После полёта стал символом мира, погиб в авиакатастрофе в 1968 году.',
    facts: { birth: '1934', death: '1968', born: 'Клушино', died: 'Киржач', country: 'СССР', job: 'космонавт, лётчик' },
    family: { 'Отец': ['Алексей Иванович Гагарин'], 'Мать': ['Анна Тимофеевна Гагарина'], 'Супруг(а)': ['Валентина Гагарина'] }
  },
  'сергей королёв': {
    title: 'Сергей Королёв',
    text: 'Сергей Павлович Королёв (1907—1966) — советский инженер и конструктор, основатель космонавтики. Родился в Житомире в семье учителя. Окончил МВТУ, создавал ракеты и ракетные двигатели. В 1957 году под его руководством был запущен первый искусственный спутник Земли, а в 1961 году — корабль с Юрием Гагариным. Королёв работал в секретных условиях и был известен только как «Главный конструктор». Его называют создателем советской космической программы.',
    facts: { birth: '1907', death: '1966', born: 'Житомир', died: 'Москва', country: 'СССР', job: 'инженер, конструктор' },
    family: { 'Отец': ['Павел Яковлевич Королёв'], 'Мать': ['Мария Николаевна Баланина'], 'Супруг(а)': ['Ксения Vincentini'] }
  },
  'дмитрий менделеев': {
    title: 'Дмитрий Менделеев',
    text: 'Дмитрий Иванович Менделеев (1834—1907) — русский химик, создатель периодической системы элементов. Родился в Тобольске в семье учителя. Окончил Петербургский университет и стал профессором. Открыл периодический закон и предсказал существование ещё не известных элементов. Его таблица до сих пор используется во всём мире.',
    facts: { birth: '1834', death: '1907', born: 'Тобольск', died: 'Санкт-Петербург', country: 'Россия', job: 'химик' },
    family: { 'Отец': ['Иван Павлович Менделеев'], 'Супруг(а)': ['Анна Добротворская'] }
  },
  'михаил ломоносов': {
    title: 'Михаил Ломоносов',
    text: 'Михаил Васильевич Ломоносов (1711—1765) — русский учёный, поэт и просветитель. Родился в деревне Мишанинская на Севере в семье крестьянина-рыболова. В 19 лет пешком ушёл в Москву учиться, затем учился в Петербурге и за границей. Работал в Академии наук, написал труды по химии, физике и астрономии, открыл закон сохранения массы. Заложил основы русского литературного языка, написал оды и трагедии. Основал Московский университет.',
    facts: { birth: '1711', death: '1765', born: 'Мишанинская', died: 'Санкт-Петербург', country: 'Россия', job: 'учёный, поэт' },
    family: { 'Отец': ['Василий Дорофеевич Ломоносов'], 'Супруг(а)': ['Елизавета Зиновьева'] }
  },
  'александр суворов': {
    title: 'Александр Суворов',
    text: 'Александр Васильевич Суворов (1730—1800) — великий русский полководец, генералиссимус. Родился в семье дворянина, с детства готовился к военной службе. Отличился в Семилетней войне, а затем в войнах с Турцией и Польшей. Командовал штурмом Измаила, перешёл через Альпы и завоевал славу непобедимого. Его книга «Наука побеждать» учила действовать быстро и смело. Суворов не проиграл ни одного сражения.',
    facts: { birth: '1730', death: '1800', born: 'Москва', died: 'Санкт-Петербург', country: 'Россия', job: 'полководец' },
    family: { 'Отец': ['Василий Иванович Суворов'], 'Супруг(а)': ['Варвара Протасова'] }
  },
  'георгий жуков': {
    title: 'Георгий Жуков',
    text: 'Георгий Константинович Жуков (1896—1974) — советский полководец, Маршал Советского Союза. Родился в деревне Стрелковка в крестьянской семье. Воевал в Гражданскую войну, командовал фронтами в годы Великой Отечественной войны. Под его руководством провели контрнаступление под Москвой, оборону Сталинграда и штурм Берлина. После войны был министром обороны, но затем попал в опалу. Считается одним из самых выдающихся полководцев XX века.',
    facts: { birth: '1896', death: '1974', born: 'Стрелковка', died: 'Москва', country: 'СССР', job: 'военачальник' },
    family: { 'Отец': ['Константин Андреевич Жуков'], 'Мать': ['Анна Григорьевна Жукова'] }
  },
  'мария кюри': {
    title: 'Мария Кюри',
    text: 'Мария Склодовская-Кюри (1867—1934) — выдающийся физик и химик, лауреат двух Нобелевских премий. Родилась в Варшаве, училась в Сорбонне в Париже. Вместе с мужем Пьером Кюри исследовала радиоактивность и открыла элементы полоний и радий. После смерти мужа продолжила работу самостоятельно. Получила Нобелевские премии по физике (1903) и химии (1911). Её исследования заложили основы ядерной физики.',
    facts: { birth: '1867', death: '1934', born: 'Варшава', died: 'Пасси', country: 'Польша, Франция', job: 'физик, химик' },
    family: { 'Супруг(а)': ['Пьер Кюри'] }
  },
  'исаак ньютон': {
    title: 'Исаак Ньютон',
    text: 'Исаак Ньютон (1643—1727) — английский физик, математик и астроном, один из величайших учёных в истории. Родился в Вулсторпе в семье фермера. Учился в Кембриджском университете. Сформулировал законы движения и закон всемирного тяготения, разработал дифференциальное исчисление, исследовал природу света и цвета. Его книга «Математические начала натуральной философии» стала основой классической механики. Был президентом Королевского общества.',
    facts: { birth: '1643', death: '1727', born: 'Вулсторп', died: 'Лондон', country: 'Англия', job: 'физик, математик' },
    family: { 'Отец': ['Исаак Ньютон-старший'], 'Мать': ['Ханна Эйскоу'] }
  }
};

function findLocal(raw, q) {
  const norm = s => String(s || '').toLowerCase().replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9\s]/gi, ' ').replace(/\s+/g, ' ').trim();
  for (const k of [norm(q), norm(raw)].filter(Boolean)) {
    if (LOCAL_PEOPLE[k]) return LOCAL_PEOPLE[k];
  }
  let best = null, bestScore = 0;
  for (const [k, v] of Object.entries(LOCAL_PEOPLE)) {
    const s = scoreCandidate(norm(q), k, v.title);
    if (s > bestScore) { bestScore = s; best = v; }
  }
  return bestScore >= 0.72 ? best : null;
}

function cleanQuery(raw) {
  let s = String(raw || '').toLowerCase().replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9\s-]/gi, ' ').replace(/\s+/g, ' ').trim();
  let words = s.split(' ').filter(w => w.length > 1 && !STOP_WORDS.has(w));
  s = words.join(' ');
  return ALIASES[s] || s || String(raw || '').trim();
}

function lev(a, b) {
  a = String(a || ''); b = String(b || '');
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let last = prev[0]; prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const old = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, last + (a[i - 1] === b[j - 1] ? 0 : 1));
      last = old;
    }
  }
  return prev[b.length];
}

function sim(a, b) {
  a = String(a || '').toLowerCase(); b = String(b || '').toLowerCase();
  if (!a || !b) return 0;
  return 1 - lev(a, b) / Math.max(a.length, b.length);
}

function scoreCandidate(q, title, desc) {
  const t = String(title || '').toLowerCase().replace(/ё/g, 'е');
  const words = q.split(' ').filter(Boolean);
  let best = 0;
  for (const w of words) {
    if (t.includes(w)) best = Math.max(best, 0.92);
    for (const tw of t.split(' ')) best = Math.max(best, sim(w, tw) * 0.9);
    best = Math.max(best, sim(w, t) * 0.85);
  }
  best = Math.max(best, sim(q, t));
  if (desc && /человек|личность|персон/i.test(desc)) best += 0.04;
  return best;
}

const WIKI_FALLBACK = 'https://ru.m.wikipedia.org/w/api.php';

async function wikiSearchRaw(name, limit = 8) {
  const params = {
    action: 'query', list: 'search', srsearch: name,
    format: 'json', origin: '*', srlimit: String(limit),
  };
  for (const endpoint of [WIKI, WIKI_FALLBACK]) {
    try {
      const data = await apiGet(endpoint, params);
      const hits = data.query?.search || [];
      if (hits.length) return hits;
    } catch (e) {}
  }
  return [];
}

async function wdSearch(name) {
  const data = await apiGet(WD_API, {
    action: 'wbsearchentities', search: name, language: 'ru',
    format: 'json', origin: '*', limit: '10',
  });
  return data.search || [];
}

async function ddgEntity(name) {
  const data = await apiGet(DDG, {
    q: name, format: 'json', no_html: 1, no_redirect: 1,
  });
  return data.Heading ? { title: data.Heading, desc: data.Abstract || '' } : null;
}

async function findPerson(raw) {
  let q = cleanQuery(raw);

  // Сначала проверяем встроенную базу — она работает без интернета.
  const local = findLocal(raw, q);
  if (local) {
    return {
      query: q,
      best: {
        title: local.title, text: local.text, facts: local.facts,
        family: local.family, source: 'local', desc: local.text.slice(0, 150),
      },
      all: [],
    };
  }

  const queries = [q];
  if (q !== String(raw).trim()) queries.push(String(raw).trim());

  let candidates = [];
  for (const query of queries) {
    const [wiki, wd] = await Promise.allSettled([wikiSearchRaw(query, 8), wdSearch(query)]);
    if (wiki.status === 'fulfilled') {
      for (const h of wiki.value) {
        candidates.push({
          title: h.title, pageid: h.pageid,
          desc: String(h.snippet || '').replace(/<[^>]+>/g, ''),
          score: scoreCandidate(q, h.title, h.snippet || ''), source: 'wikipedia',
        });
      }
    }
    if (wd.status === 'fulfilled') {
      for (const e of wd.value) {
        candidates.push({
          title: e.label, qid: e.id, desc: e.description || '',
          score: scoreCandidate(q, e.label + ' ' + (e.aliases || []).join(' '), e.description || ''),
          source: 'wikidata',
        });
      }
    }
    if (candidates.length) break;

    // Если Википедия/Викиданные ничего не нашли — пробуем «яндексоподобный» ответ.
    try {
      const dd = await ddgEntity(query);
      if (dd) {
        const more = await wikiSearchRaw(dd.title, 3);
        for (const h of more) {
          candidates.push({
            title: h.title, pageid: h.pageid,
            desc: dd.desc || String(h.snippet || '').replace(/<[^>]+>/g, ''),
            score: scoreCandidate(q, h.title, dd.desc || ''), source: 'duckduckgo',
          });
        }
      }
    } catch (e) {}
    if (candidates.length) break;
  }

  const seen = new Set();
  candidates = candidates.filter(c => {
    const k = String(c.title || '').toLowerCase().replace(/ё/g, 'е');
    if (!k || seen.has(k)) return false;
    seen.add(k); return true;
  }).sort((a, b) => b.score - a.score);

  return { query: q, best: candidates[0] || null, all: candidates.slice(0, 5) };
}

async function wikiArticle(pageid) {
  const params = {
    action: 'query', prop: 'extracts|pageprops', pageids: String(pageid),
    explaintext: '1', format: 'json', origin: '*',
  };
  let data = null;
  for (const endpoint of [WIKI, WIKI_FALLBACK]) {
    try {
      data = await apiGet(endpoint, params);
      if (data.query?.pages?.[String(pageid)]) break;
    } catch (e) {}
  }
  data ||= {};
  const page = data.query?.pages?.[String(pageid)] || {};
  return {
    title: page.title || '',
    text: page.extract || '',
    qid: page.pageprops?.wikibase_item || '',
  };
}

async function sparql(q) {
  const data = await apiGet(SPARQL, { query: q, format: 'json' });
  return data.results?.bindings || [];
}

const val = (row, key) => row?.[key]?.value || '';

async function collect(raw) {
  const found = await findPerson(raw);
  if (!found.best) return null;

  let best = found.best;

  if (best.source === 'local') {
    return {
      title: best.title,
      text: best.text,
      qid: '',
      facts: best.facts,
      family: best.family,
      query: found.query,
      matched: best.title,
      candidates: [],
      wikiUrl: 'https://ru.wikipedia.org/wiki/' + encodeURIComponent(best.title.replace(/ /g, '_')),
      wdUrl: '',
    };
  }

  let pageid = best.pageid;

  if (!pageid && best.qid) {
    try {
      const data = await apiGet(WD_API, {
        action: 'wbgetentities', ids: best.qid, props: 'sitelinks',
        sitefilter: 'ruwiki', format: 'json', origin: '*',
      });
      const ruTitle = data.entities?.[best.qid]?.sitelinks?.ruwiki?.title;
      if (ruTitle) {
        const wikiHits = await wikiSearchRaw(ruTitle, 3);
        pageid = wikiHits[0]?.pageid;
      }
    } catch (e) {}
    if (!pageid) {
      try {
        const wikiHits = await wikiSearchRaw(best.title, 3);
        pageid = wikiHits[0]?.pageid;
      } catch (e) {}
    }
  }
  if (!pageid) return null;

  let article;
  try {
    article = await wikiArticle(pageid);
  } catch (e) {
    // Если статья временно недоступна — всё равно сделаем доклад из найденного.
    article = { title: best.title, text: best.desc || '', qid: best.qid || '' };
  }
  const qid = article.qid || best.qid || '';
  let facts = {}, family = {};

  if (qid) {
    try {
      const fRows = await sparql(`
        SELECT ?dob ?dod ?bornLabel ?diedLabel ?countryLabel ?jobLabel WHERE {
          VALUES ?person { wd:${qid} }
          OPTIONAL { ?person wdt:P569 ?dob }
          OPTIONAL { ?person wdt:P570 ?dod }
          OPTIONAL { ?person wdt:P19 ?born }
          OPTIONAL { ?person wdt:P20 ?died }
          OPTIONAL { ?person wdt:P27 ?country }
          OPTIONAL { ?person wdt:P106 ?job }
          SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
        } LIMIT 1`);
      const r = fRows[0] || {};
      facts = {
        birth: (val(r, 'dob') || '—').slice(0, 4),
        death: (val(r, 'dod') || '—').slice(0, 4),
        born: val(r, 'bornLabel') || '—',
        died: val(r, 'diedLabel') || '—',
        country: val(r, 'countryLabel') || '—',
        job: val(r, 'jobLabel') || '—',
      };
    } catch (e) {}

    try {
      const famRows = await sparql(`
        SELECT ?kind ?relLabel WHERE {
          VALUES ?person { wd:${qid} }
          VALUES (?p ?kind) {
            (wdt:P22 "Отец") (wdt:P25 "Мать") (wdt:P26 "Супруг(а)") (wdt:P40 "Дети")
          }
          ?person ?p ?rel .
          SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
        } ORDER BY ?kind`);
      for (const row of famRows) {
        const k = val(row, 'kind');
        (family[k] ||= []).push(val(row, 'relLabel'));
      }
    } catch (e) {}
  }

  return {
    ...article, facts, family,
    query: found.query, matched: article.title, candidates: found.all,
    wikiUrl: 'https://ru.wikipedia.org/wiki/' + article.title.replace(/ /g, '_'),
    wdUrl: qid ? `https://www.wikidata.org/wiki/${qid}` : '',
  };
}

function sentences(text, limit = 6) {
  return String(text).replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ---------- Текст доклада ----------

function makeReport(d, note) {
  const f = d.facts;
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. Родился(ась) в ${f.birth} году в городе ${f.born}. Гражданство: ${f.country}. Род деятельности: ${f.job}.`);
  lines.push('', 'Основная часть');
  for (const s of sentences(d.text, 6)) lines.push(s);
  lines.push('', 'Семья');
  const fam = Object.entries(d.family);
  if (fam.length) for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}`);
  else lines.push('Сведения о семье в открытых источниках отсутствуют.');
  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);
  if (note) lines.push('', 'Дополнительно учесть: ' + note);
  lines.push('', 'Источники:', d.wikiUrl);
  if (d.wdUrl) lines.push(d.wdUrl);
  return lines.join('\n');
}

// ---------- Презентация ----------

async function makePptx(d, styleKey, colorKey, filePath) {
  const color = COLORS[colorKey][1];
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  function addSlide(title, bullets, fontSize = 20) {
    const slide = pptx.addSlide();
    slide.background = { color };
    slide.addText(title, {
      x: 0.7, y: 0.5, w: 12, h: 1, fontSize: styleKey === 'kids' ? 38 : 34,
      bold: true, color: 'FFFFFF', align: 'center',
    });
    const text = bullets.map(b => '• ' + b).join('\n');
    slide.addText(text, {
      x: 0.9, y: 1.8, w: 11.6, h: 5.2,
      fontSize: styleKey === 'kids' ? fontSize + 2 : fontSize,
      color: 'FFFFFF', breakLine: false, paraSpaceAfter: 10,
    });
  }

  const f = d.facts;
  addSlide(d.title, ['Школьная презентация', `Стиль: ${STYLES[styleKey]}`, `Цвет: ${COLORS[colorKey][0]}`], 24);
  addSlide('Кто это?', [
    `Родился(ась): ${f.birth}`,
    `Место рождения: ${f.born}`,
    `Гражданство: ${f.country}`,
    `Деятельность: ${f.job}`,
  ]);
  addSlide('Биография', sentences(d.text, 4), 17);
  const fam = Object.entries(d.family).map(([k, v]) => `${k}: ${v.join(', ')}`);
  addSlide('Семья', fam.length ? fam : ['Сведений нет']);
  addSlide('География', [
    `Рождение: ${f.born}`, `Смерть: ${f.died}`, `Страна: ${f.country}`,
  ]);
  addSlide('Интересные факты', sentences(d.text, 3).reverse(), 17);
  addSlide('Вывод', [
    `${d.title} — важная историческая личность.`,
    'Его/её жизнь связана с ключевыми событиями эпохи.',
    'Изучение биографии помогает лучше понять историю.',
  ]);
  addSlide('Источники', [d.wikiUrl, d.wdUrl || 'Викиданные'].filter(Boolean), 16);

  await pptx.writeFile({ fileName: filePath });
}

// ---------- Хендлеры ----------

function matchesWord(text, words) {
  const low = String(text || '').toLowerCase().replace(/ё/g, 'е').trim();
  return words.some(w => low === w || low.startsWith(w + ' ') || low.includes(' ' + w + ' '));
}

async function askStyle(ctx, s) {
  s.step = 'pres_style';
  await ctx.reply('Выбери стиль презентации:', styleKeyboard());
}

async function startPresentationForLast(ctx, s) {
  if (s.lastName) {
    s.name = s.lastName;
    s.note = s.lastNote || '';
    await ctx.reply(`Делаю презентацию про: ${s.lastName}`);
    await askStyle(ctx, s);
  } else {
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
  }
}

async function handleTextRequest(ctx, s, text) {
  const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
  if (!parts.length) return ctx.reply('Напиши, про кого сделать работу.');
  await ctx.reply('Понял запрос. Ищу и сверяю по нескольким источникам...');
  try {
    const d = await collect(parts[0]);
    if (!d) {
      return ctx.reply(
        'Не нашёл этого человека. Википедия сейчас может ограничивать запросы с нашего сервера. ' +
        'Напиши полное имя, например: Альберт Эйнштейн, Александр Пушкин, Юрий Гагарин. ' +
        'Если нужного имени нет во встроенной базе — попробуй через пару минут.'
      );
    }
    if (d.matched && d.query !== d.matched.toLowerCase()) {
      await ctx.reply(`Похоже, ты имел(а) в виду: ${d.matched}`);
    }
    const note = parts.slice(1).join(' ');
    const report = makeReport(d, note);
    for (let i = 0; i < report.length; i += 3500) {
      await ctx.reply(report.slice(i, i + 3500));
    }
    s.lastName = d.matched || parts[0];
    s.lastNote = note;
    s.awaitingPres = true;
    s.step = 'menu';
    await ctx.reply('Хочешь также презентацию?', yesNoKeyboard());
  } catch (e) {
    console.error(e);
    const msg = String(e.message || '');
    if (/429/.test(msg)) {
      await ctx.reply('Википедия временно ограничила число запросов (429). ' +
        'Я уже повторил несколько раз — попробуй ещё раз через минуту.');
    } else {
      await ctx.reply('Ошибка: ' + msg);
    }
  }
}

bot.start(async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply(
    'Привет! Я помогу с школьной работой про человека.\n\nВыбери, что нужно сделать:',
    mainMenu
  );
});

bot.command('help', async (ctx) => {
  await ctx.reply(
    '📝 Готовый текст — напиши, про кого, получишь текст для вставки.\n' +
    '📊 Презентация — напиши тему, ответь на вопросы про дизайн и получи файл.\n\n' +
    '/start — начать заново'
  );
});

bot.command('cancel', async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply('Отменено. /start — начать заново.', mainMenu);
});

bot.on('text', async (ctx) => {
  const text = ctx.message.text.trim();
  const s = session(ctx);

  if (s.awaitingPres && s.step === 'menu') {
    if (matchesWord(text, YES_WORDS)) {
      s.awaitingPres = false;
      await startPresentationForLast(ctx, s);
      return;
    }
    if (matchesWord(text, NO_WORDS)) {
      s.awaitingPres = false;
      await ctx.reply('Хорошо. Если передумаешь — нажми 📊 Презентация.', mainMenu);
      return;
    }
    s.awaitingPres = false;
  }

  if (text === TEXT_BTN) {
    s.step = 'text_topic';
    await ctx.reply(
      'Напиши, про кого сделать текст, и что учесть.\nНапример:\nУинстон Черчилль\nДля 7 класса, про семью и войну',
      removeKb
    );
    return;
  }

  if (text === PRES_BTN) {
    s.awaitingPres = false;
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
    if (!parts.length) return ctx.reply('Напиши имя человека.');
    s.name = parts[0];
    s.note = parts.slice(1).join(' ');
    s.step = 'pres_style';
    await ctx.reply('Выбери стиль презентации:', styleKeyboard());
    return;
  }

  // Если человек просто написал «напиши про энштейна» — сразу понимаем это как запрос текста.
  if (s.step === 'menu') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  await ctx.reply('Выбери кнопку или напиши /start.', mainMenu);
});

bot.action('pres:yes', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Отлично, делаю презентацию.');
  s.awaitingPres = false;
  await startPresentationForLast(ctx, s);
});

bot.action('pres:no', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Хорошо. Если передумаешь — нажми 📊 Презентация.');
  s.awaitingPres = false;
  s.step = 'menu';
  await ctx.reply('Что ещё сделать?', mainMenu);
});

bot.action(/^style:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  s.style = key;
  s.step = 'pres_color';
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Стиль: ${STYLES[key]}. Теперь выбери цвет:`);
  await ctx.reply('Цвет:', colorKeyboard());
});

bot.action(/^color:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Цвет: ${COLORS[key][0]}. Делаю презентацию...`);
  try {
    const d = await collect(s.name);
    if (!d) return ctx.reply('Не нашёл человека в Википедии.');
    const filePath = path.join(os.tmpdir(), `presentation_${ctx.chat.id}.pptx`);
    await makePptx(d, s.style, key, filePath);
    await ctx.replyWithDocument(
      { source: fs.createReadStream(filePath), filename: `Презентация_${d.title}.pptx` },
      { caption: `Готовая презентация: ${d.title}\nСтиль: ${STYLES[s.style]}, цвет: ${COLORS[key][0]}` }
    );
    fs.unlinkSync(filePath);
    await ctx.reply('Ещё что-нибудь?', mainMenu);
    s.step = 'menu';
  } catch (e) {
    console.error(e);
    await ctx.reply('Не удалось сделать презентацию: ' + e.message);
    await ctx.reply('/start — вернуться в меню', mainMenu);
    s.step = 'menu';
  }
});

async function launchWithRetry() {
  for (let i = 1; i <= 12; i++) {
    try {
      await bot.launch();
      console.log('Бот запущен (polling)');
      return;
    } catch (e) {
      console.log(`Попытка ${i}/12 не удалась: ${e.message}. Повтор через 5 сек...`);
      await new Promise(r => setTimeout(r, 5000));
    }
  }
  console.error('Не удалось подключиться к Telegram. Проверь, что Happ подключён и прокси 10808 работает.');
  process.exit(1);
}

// --- Облако: webhook вместо polling ---
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.RENDER_EXTERNAL_URL
  ? process.env.RENDER_EXTERNAL_URL.replace(/\/$/, '') + '/webhook'
  : (process.env.WEBHOOK_URL || '');

async function start() {
  if (WEBHOOK_URL) {
    const app = express();
    app.use(express.json({ limit: '1mb' }));
    app.get('/', (req, res) => res.send('Bot is alive'));
    app.post('/webhook', bot.webhookCallback('/webhook'));

    await bot.telegram.setWebhook(WEBHOOK_URL);
    app.listen(PORT, () => {
      console.log('Webhook server started on port ' + PORT);
      console.log('Webhook URL: ' + WEBHOOK_URL);
    });
    return;
  }
  await launchWithRetry();
}

start();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));


}

// Happ использует локальный SOCKS-прокси 127.0.0.1:10808 (см. config.json).
// Если у тебя другой прокси — задай его переменной BOT_PROXY в run_bot.bat.
// На Windows используем локальный прокси Happ. На Linux-хостинге прокси не нужен.
const HAPP_PROXY = process.platform === 'win32' ? 'socks://127.0.0.1:10808' : '';
const PROXY = process.env.BOT_PROXY || HAPP_PROXY || detectProxy();
let agent;
if (PROXY) {
  agent = /^socks/i.test(PROXY) ? new SocksProxyAgent(PROXY) : new HttpsProxyAgent(PROXY);
  console.log('Использую прокси: ' + PROXY);
} else {
  console.log('Прокси не найден — Trying direct connection.');
}
const netOpts = agent ? { agent, httpsAgent: agent } : {};

const bot = new Telegraf(TOKEN, agent ? { telegram: { agent } } : {});
const WIKI = 'https://ru.wikipedia.org/w/api.php';
const WD_API = 'https://www.wikidata.org/w/api.php';
const DDG = 'https://api.duckduckgo.com/';
const SPARQL = 'https://query.wikidata.org/sparql';

// Википедия отклоняет запросы без нормального User-Agent — из-за этого был 403.
const UA = 'SchoolBot/1.0 (school project; t.me/schol_spg_bot)';
axios.defaults.headers.common['User-Agent'] = UA;
axios.defaults.headers.common['Accept'] = 'application/json';

// --- Защита от лимитов Википедии/Викиданных (ошибка 429) ---
// Render использует общий IP, поэтому запросы идут реже, повторяются при 429
// и кэшируются, чтобы повторный вопрос про того же человека не тратил лимит.
const apiCache = new Map();
let lastApiCall = 0;
let apiBusy = false;

async function apiGet(url, params) {
  const key = url + '|' + JSON.stringify(params || {});
  if (apiCache.has(key)) return apiCache.get(key);

  // Несколько вызовов подряд (поиск в Википедии + Викиданных) идут строго по очереди.
  while (apiBusy) await new Promise(r => setTimeout(r, 120));
  apiBusy = true;
  try {
    const gap = 400 - (Date.now() - lastApiCall);
    if (gap > 0) await new Promise(r => setTimeout(r, gap));

    let lastErr;
    for (let attempt = 0; attempt < 5; attempt++) {
      lastApiCall = Date.now();
      try {
        const { data } = await axios.get(url, { params, timeout: 30000, ...netOpts });
        apiCache.set(key, data);
        return data;
      } catch (e) {
        lastErr = e;
        const status = e.response?.status;
        const retryable = status === 429 || status === 500 || status === 502 ||
                          status === 503 || status === 504 || !status;
        if (!retryable || attempt === 4) break;
        const header = Number(e.response?.headers?.['retry-after']);
        const wait = Number.isFinite(header) && header > 0
          ? header * 1000
          : Math.min(2000 * Math.pow(2, attempt), 15000);
        await new Promise(r => setTimeout(r, wait));
      }
    }
    throw lastErr;
  } finally {
    apiBusy = false;
  }
}

const TEXT_BTN = '📝 Готовый текст';
const PRES_BTN = '📊 Презентация';

const mainMenu = Markup.keyboard([[TEXT_BTN, PRES_BTN]]).resize();
const removeKb = Markup.removeKeyboard();

const YES_WORDS = ['да', 'давай', 'давай же', 'ага', 'угу', 'конечно', 'ок', 'окей',
  'ладно', 'хочу', 'сделай', 'презентацию', 'презентация', 'yes', 'yeah', 'добре'];
const NO_WORDS = ['нет', 'не надо', 'не хочу', 'не сейчас', 'позже', 'стоп', 'отмена', 'no'];

const yesNoKeyboard = () => Markup.inlineKeyboard([
  [Markup.button.callback('Да, сделать презентацию', 'pres:yes'),
   Markup.button.callback('Нет, спасибо', 'pres:no')],
]);

const STYLES = { classic: 'Классический', modern: 'Современный', kids: 'Детский' };
const COLORS = {
  blue: ['Синий', '1F4E8C'],
  green: ['Зелёный', '226E4E'],
  red: ['Красный', '962D37'],
  violet: ['Фиолетовый', '5E378C'],
};

const styleKeyboard = () => Markup.inlineKeyboard(
  Object.entries(STYLES).map(([k, label]) => Markup.button.callback(label, `style:${k}`))
);
const colorKeyboard = () => Markup.inlineKeyboard(
  Object.entries(COLORS).map(([k, v]) => Markup.button.callback(v[0], `color:${k}`))
);

// Простейшие сессии в памяти
const sessions = new Map();
function session(ctx) {
  const id = ctx.chat.id;
  if (!sessions.has(id)) sessions.set(id, { step: 'menu' });
  return sessions.get(id);
}

// ---------- Открытые данные ----------

// --- Понимание разговорных запросов ("напиши про энштейна" и т.п.) ---
const STOP_WORDS = new Set([
  'напиши','написать','напишите','про','об','about','сделай','сделать','сделайте',
  'давай','дай','пришли','инфа','инфу','информация','информацию',
  'текст','доклад','доклада','презентация','презентацию','презентации',
  'биография','биографию','биографии','расскажи','рассказать','кто','такой','такая',
  'что','какой','какая','нужно','надо','помоги','найди','ищи',
  'поищи','сведения','данные','материал','материалы','тему','тема','для','класс',
  'класса','школе','школа','школь','человек','личность','жизнь','жизни'
]);

// Частые сокращения и опечатки -> полное имя.
const ALIASES = {
  'энштейн': 'альберт эйнштейн', 'энштейна': 'альберт эйнштейн', 'айнштейн': 'альберт эйнштейн',
  'пушкин': 'александр пушкин', 'пушкина': 'александр пушкин',
  'лермонтов': 'михаил лермонтов', 'толстой': 'лев толстой', 'чехов': 'антон чехов',
  'достоевский': 'фёдор достоевский', 'гоголь': 'николай гоголь',
  'черчилль': 'уинстон черчилль', 'наполеон': 'наполеон бонапарт',
  'сталин': 'иосиф сталин', 'ленин': 'владимир ленин', 'гитлер': 'адольф гитлер',
  'петр 1': 'пётр великий', 'петр великий': 'пётр великий', 'пётр 1': 'пётр великий',
  'екатерина 2': 'екатерина великая', 'екатерина великая': 'екатерина великая',
  'гагарин': 'юрий гагарин', 'королев': 'сергей королёв', 'менделеев': 'дмитрий менделеев',
  'ломоносов': 'михаил ломоносов', 'суворов': 'александр суворов', 'жуков': 'георгий жуков'
};

function cleanQuery(raw) {
  let s = String(raw || '').toLowerCase().replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9\s-]/gi, ' ').replace(/\s+/g, ' ').trim();
  let words = s.split(' ').filter(w => w.length > 1 && !STOP_WORDS.has(w));
  s = words.join(' ');
  return ALIASES[s] || s || String(raw || '').trim();
}

function lev(a, b) {
  a = String(a || ''); b = String(b || '');
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let last = prev[0]; prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const old = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, last + (a[i - 1] === b[j - 1] ? 0 : 1));
      last = old;
    }
  }
  return prev[b.length];
}

function sim(a, b) {
  a = String(a || '').toLowerCase(); b = String(b || '').toLowerCase();
  if (!a || !b) return 0;
  return 1 - lev(a, b) / Math.max(a.length, b.length);
}

function scoreCandidate(q, title, desc) {
  const t = String(title || '').toLowerCase().replace(/ё/g, 'е');
  const words = q.split(' ').filter(Boolean);
  let best = 0;
  for (const w of words) {
    if (t.includes(w)) best = Math.max(best, 0.92);
    for (const tw of t.split(' ')) best = Math.max(best, sim(w, tw) * 0.9);
    best = Math.max(best, sim(w, t) * 0.85);
  }
  best = Math.max(best, sim(q, t));
  if (desc && /человек|личность|персон/i.test(desc)) best += 0.04;
  return best;
}

async function wikiSearchRaw(name, limit = 8) {
  const data = await apiGet(WIKI, {
    action: 'query', list: 'search', srsearch: name,
    format: 'json', origin: '*', srlimit: String(limit),
  });
  return data.query?.search || [];
}

async function wdSearch(name) {
  const data = await apiGet(WD_API, {
    action: 'wbsearchentities', search: name, language: 'ru',
    format: 'json', origin: '*', limit: '10',
  });
  return data.search || [];
}

async function ddgEntity(name) {
  const data = await apiGet(DDG, {
    q: name, format: 'json', no_html: 1, no_redirect: 1,
  });
  return data.Heading ? { title: data.Heading, desc: data.Abstract || '' } : null;
}

async function findPerson(raw) {
  let q = cleanQuery(raw);
  const queries = [q];
  if (q !== String(raw).trim()) queries.push(String(raw).trim());

  let candidates = [];
  for (const query of queries) {
    const [wiki, wd] = await Promise.allSettled([wikiSearchRaw(query, 8), wdSearch(query)]);
    if (wiki.status === 'fulfilled') {
      for (const h of wiki.value) {
        candidates.push({
          title: h.title, pageid: h.pageid,
          desc: String(h.snippet || '').replace(/<[^>]+>/g, ''),
          score: scoreCandidate(q, h.title, h.snippet || ''), source: 'wikipedia',
        });
      }
    }
    if (wd.status === 'fulfilled') {
      for (const e of wd.value) {
        candidates.push({
          title: e.label, qid: e.id, desc: e.description || '',
          score: scoreCandidate(q, e.label + ' ' + (e.aliases || []).join(' '), e.description || ''),
          source: 'wikidata',
        });
      }
    }
    if (candidates.length) break;

    // Если Википедия/Викиданные ничего не нашли — пробуем «яндексоподобный» ответ.
    try {
      const dd = await ddgEntity(query);
      if (dd) {
        const more = await wikiSearchRaw(dd.title, 3);
        for (const h of more) {
          candidates.push({
            title: h.title, pageid: h.pageid,
            desc: dd.desc || String(h.snippet || '').replace(/<[^>]+>/g, ''),
            score: scoreCandidate(q, h.title, dd.desc || ''), source: 'duckduckgo',
          });
        }
      }
    } catch (e) {}
    if (candidates.length) break;
  }

  const seen = new Set();
  candidates = candidates.filter(c => {
    const k = String(c.title || '').toLowerCase().replace(/ё/g, 'е');
    if (!k || seen.has(k)) return false;
    seen.add(k); return true;
  }).sort((a, b) => b.score - a.score);

  return { query: q, best: candidates[0] || null, all: candidates.slice(0, 5) };
}

async function wikiArticle(pageid) {
  const data = await apiGet(WIKI, {
    action: 'query', prop: 'extracts|pageprops', pageids: String(pageid),
    explaintext: '1', format: 'json', origin: '*',
  });
  const page = data.query?.pages?.[String(pageid)] || {};
  return {
    title: page.title || '',
    text: page.extract || '',
    qid: page.pageprops?.wikibase_item || '',
  };
}

async function sparql(q) {
  const data = await apiGet(SPARQL, { query: q, format: 'json' });
  return data.results?.bindings || [];
}

const val = (row, key) => row?.[key]?.value || '';

async function collect(raw) {
  const found = await findPerson(raw);
  if (!found.best) return null;

  let best = found.best;
  let pageid = best.pageid;

  if (!pageid && best.qid) {
    try {
      const data = await apiGet(WD_API, {
        action: 'wbgetentities', ids: best.qid, props: 'sitelinks',
        sitefilter: 'ruwiki', format: 'json', origin: '*',
      });
      const ruTitle = data.entities?.[best.qid]?.sitelinks?.ruwiki?.title;
      if (ruTitle) {
        const wikiHits = await wikiSearchRaw(ruTitle, 3);
        pageid = wikiHits[0]?.pageid;
      }
    } catch (e) {}
    if (!pageid) {
      try {
        const wikiHits = await wikiSearchRaw(best.title, 3);
        pageid = wikiHits[0]?.pageid;
      } catch (e) {}
    }
  }
  if (!pageid) return null;

  let article;
  try {
    article = await wikiArticle(pageid);
  } catch (e) {
    // Если статья временно недоступна — всё равно сделаем доклад из найденного.
    article = { title: best.title, text: best.desc || '', qid: best.qid || '' };
  }
  const qid = article.qid || best.qid || '';
  let facts = {}, family = {};

  if (qid) {
    try {
      const fRows = await sparql(`
        SELECT ?dob ?dod ?bornLabel ?diedLabel ?countryLabel ?jobLabel WHERE {
          VALUES ?person { wd:${qid} }
          OPTIONAL { ?person wdt:P569 ?dob }
          OPTIONAL { ?person wdt:P570 ?dod }
          OPTIONAL { ?person wdt:P19 ?born }
          OPTIONAL { ?person wdt:P20 ?died }
          OPTIONAL { ?person wdt:P27 ?country }
          OPTIONAL { ?person wdt:P106 ?job }
          SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
        } LIMIT 1`);
      const r = fRows[0] || {};
      facts = {
        birth: (val(r, 'dob') || '—').slice(0, 4),
        death: (val(r, 'dod') || '—').slice(0, 4),
        born: val(r, 'bornLabel') || '—',
        died: val(r, 'diedLabel') || '—',
        country: val(r, 'countryLabel') || '—',
        job: val(r, 'jobLabel') || '—',
      };
    } catch (e) {}

    try {
      const famRows = await sparql(`
        SELECT ?kind ?relLabel WHERE {
          VALUES ?person { wd:${qid} }
          VALUES (?p ?kind) {
            (wdt:P22 "Отец") (wdt:P25 "Мать") (wdt:P26 "Супруг(а)") (wdt:P40 "Дети")
          }
          ?person ?p ?rel .
          SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
        } ORDER BY ?kind`);
      for (const row of famRows) {
        const k = val(row, 'kind');
        (family[k] ||= []).push(val(row, 'relLabel'));
      }
    } catch (e) {}
  }

  return {
    ...article, facts, family,
    query: found.query, matched: article.title, candidates: found.all,
    wikiUrl: 'https://ru.wikipedia.org/wiki/' + article.title.replace(/ /g, '_'),
    wdUrl: qid ? `https://www.wikidata.org/wiki/${qid}` : '',
  };
}

function sentences(text, limit = 6) {
  return String(text).replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ---------- Текст доклада ----------

function makeReport(d, note) {
  const f = d.facts;
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. Родился(ась) в ${f.birth} году в городе ${f.born}. Гражданство: ${f.country}. Род деятельности: ${f.job}.`);
  lines.push('', 'Основная часть');
  for (const s of sentences(d.text, 6)) lines.push(s);
  lines.push('', 'Семья');
  const fam = Object.entries(d.family);
  if (fam.length) for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}`);
  else lines.push('Сведения о семье в открытых источниках отсутствуют.');
  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);
  if (note) lines.push('', 'Дополнительно учесть: ' + note);
  lines.push('', 'Источники:', d.wikiUrl);
  if (d.wdUrl) lines.push(d.wdUrl);
  return lines.join('\n');
}

// ---------- Презентация ----------

async function makePptx(d, styleKey, colorKey, filePath) {
  const color = COLORS[colorKey][1];
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  function addSlide(title, bullets, fontSize = 20) {
    const slide = pptx.addSlide();
    slide.background = { color };
    slide.addText(title, {
      x: 0.7, y: 0.5, w: 12, h: 1, fontSize: styleKey === 'kids' ? 38 : 34,
      bold: true, color: 'FFFFFF', align: 'center',
    });
    const text = bullets.map(b => '• ' + b).join('\n');
    slide.addText(text, {
      x: 0.9, y: 1.8, w: 11.6, h: 5.2,
      fontSize: styleKey === 'kids' ? fontSize + 2 : fontSize,
      color: 'FFFFFF', breakLine: false, paraSpaceAfter: 10,
    });
  }

  const f = d.facts;
  addSlide(d.title, ['Школьная презентация', `Стиль: ${STYLES[styleKey]}`, `Цвет: ${COLORS[colorKey][0]}`], 24);
  addSlide('Кто это?', [
    `Родился(ась): ${f.birth}`,
    `Место рождения: ${f.born}`,
    `Гражданство: ${f.country}`,
    `Деятельность: ${f.job}`,
  ]);
  addSlide('Биография', sentences(d.text, 4), 17);
  const fam = Object.entries(d.family).map(([k, v]) => `${k}: ${v.join(', ')}`);
  addSlide('Семья', fam.length ? fam : ['Сведений нет']);
  addSlide('География', [
    `Рождение: ${f.born}`, `Смерть: ${f.died}`, `Страна: ${f.country}`,
  ]);
  addSlide('Интересные факты', sentences(d.text, 3).reverse(), 17);
  addSlide('Вывод', [
    `${d.title} — важная историческая личность.`,
    'Его/её жизнь связана с ключевыми событиями эпохи.',
    'Изучение биографии помогает лучше понять историю.',
  ]);
  addSlide('Источники', [d.wikiUrl, d.wdUrl || 'Викиданные'].filter(Boolean), 16);

  await pptx.writeFile({ fileName: filePath });
}

// ---------- Хендлеры ----------

function matchesWord(text, words) {
  const low = String(text || '').toLowerCase().replace(/ё/g, 'е').trim();
  return words.some(w => low === w || low.startsWith(w + ' ') || low.includes(' ' + w + ' '));
}

async function askStyle(ctx, s) {
  s.step = 'pres_style';
  await ctx.reply('Выбери стиль презентации:', styleKeyboard());
}

async function startPresentationForLast(ctx, s) {
  if (s.lastName) {
    s.name = s.lastName;
    s.note = s.lastNote || '';
    await ctx.reply(`Делаю презентацию про: ${s.lastName}`);
    await askStyle(ctx, s);
  } else {
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
  }
}

async function handleTextRequest(ctx, s, text) {
  const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
  if (!parts.length) return ctx.reply('Напиши, про кого сделать работу.');
  await ctx.reply('Понял запрос. Ищу и сверяю по нескольким источникам...');
  try {
    const d = await collect(parts[0]);
    if (!d) {
      return ctx.reply(
        'Не смог понять, о ком речь. Напиши имя и фамилию полностью, ' +
        'например: Альберт Эйнштейн. Или просто: напиши про Эйнштейна.'
      );
    }
    if (d.matched && d.query !== d.matched.toLowerCase()) {
      await ctx.reply(`Похоже, ты имел(а) в виду: ${d.matched}`);
    }
    const note = parts.slice(1).join(' ');
    const report = makeReport(d, note);
    for (let i = 0; i < report.length; i += 3500) {
      await ctx.reply(report.slice(i, i + 3500));
    }
    s.lastName = d.matched || parts[0];
    s.lastNote = note;
    s.awaitingPres = true;
    s.step = 'menu';
    await ctx.reply('Хочешь также презентацию?', yesNoKeyboard());
  } catch (e) {
    console.error(e);
    const msg = String(e.message || '');
    if (/429/.test(msg)) {
      await ctx.reply('Википедия временно ограничила число запросов (429). ' +
        'Я уже повторил несколько раз — попробуй ещё раз через минуту.');
    } else {
      await ctx.reply('Ошибка: ' + msg);
    }
  }
}

bot.start(async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply(
    'Привет! Я помогу с школьной работой про человека.\n\nВыбери, что нужно сделать:',
    mainMenu
  );
});

bot.command('help', async (ctx) => {
  await ctx.reply(
    '📝 Готовый текст — напиши, про кого, получишь текст для вставки.\n' +
    '📊 Презентация — напиши тему, ответь на вопросы про дизайн и получи файл.\n\n' +
    '/start — начать заново'
  );
});

bot.command('cancel', async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply('Отменено. /start — начать заново.', mainMenu);
});

bot.on('text', async (ctx) => {
  const text = ctx.message.text.trim();
  const s = session(ctx);

  if (s.awaitingPres && s.step === 'menu') {
    if (matchesWord(text, YES_WORDS)) {
      s.awaitingPres = false;
      await startPresentationForLast(ctx, s);
      return;
    }
    if (matchesWord(text, NO_WORDS)) {
      s.awaitingPres = false;
      await ctx.reply('Хорошо. Если передумаешь — нажми 📊 Презентация.', mainMenu);
      return;
    }
    s.awaitingPres = false;
  }

  if (text === TEXT_BTN) {
    s.step = 'text_topic';
    await ctx.reply(
      'Напиши, про кого сделать текст, и что учесть.\nНапример:\nУинстон Черчилль\nДля 7 класса, про семью и войну',
      removeKb
    );
    return;
  }

  if (text === PRES_BTN) {
    s.awaitingPres = false;
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
    if (!parts.length) return ctx.reply('Напиши имя человека.');
    s.name = parts[0];
    s.note = parts.slice(1).join(' ');
    s.step = 'pres_style';
    await ctx.reply('Выбери стиль презентации:', styleKeyboard());
    return;
  }

  // Если человек просто написал «напиши про энштейна» — сразу понимаем это как запрос текста.
  if (s.step === 'menu') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  await ctx.reply('Выбери кнопку или напиши /start.', mainMenu);
});

bot.action('pres:yes', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Отлично, делаю презентацию.');
  s.awaitingPres = false;
  await startPresentationForLast(ctx, s);
});

bot.action('pres:no', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Хорошо. Если передумаешь — нажми 📊 Презентация.');
  s.awaitingPres = false;
  s.step = 'menu';
  await ctx.reply('Что ещё сделать?', mainMenu);
});

bot.action(/^style:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  s.style = key;
  s.step = 'pres_color';
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Стиль: ${STYLES[key]}. Теперь выбери цвет:`);
  await ctx.reply('Цвет:', colorKeyboard());
});

bot.action(/^color:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Цвет: ${COLORS[key][0]}. Делаю презентацию...`);
  try {
    const d = await collect(s.name);
    if (!d) return ctx.reply('Не нашёл человека в Википедии.');
    const filePath = path.join(os.tmpdir(), `presentation_${ctx.chat.id}.pptx`);
    await makePptx(d, s.style, key, filePath);
    await ctx.replyWithDocument(
      { source: fs.createReadStream(filePath), filename: `Презентация_${d.title}.pptx` },
      { caption: `Готовая презентация: ${d.title}\nСтиль: ${STYLES[s.style]}, цвет: ${COLORS[key][0]}` }
    );
    fs.unlinkSync(filePath);
    await ctx.reply('Ещё что-нибудь?', mainMenu);
    s.step = 'menu';
  } catch (e) {
    console.error(e);
    await ctx.reply('Не удалось сделать презентацию: ' + e.message);
    await ctx.reply('/start — вернуться в меню', mainMenu);
    s.step = 'menu';
  }
});

async function launchWithRetry() {
  for (let i = 1; i <= 12; i++) {
    try {
      await bot.launch();
      console.log('Бот запущен (polling)');
      return;
    } catch (e) {
      console.log(`Попытка ${i}/12 не удалась: ${e.message}. Повтор через 5 сек...`);
      await new Promise(r => setTimeout(r, 5000));
    }
  }
  console.error('Не удалось подключиться к Telegram. Проверь, что Happ подключён и прокси 10808 работает.');
  process.exit(1);
}

// --- Облако: webhook вместо polling ---
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.RENDER_EXTERNAL_URL
  ? process.env.RENDER_EXTERNAL_URL.replace(/\/$/, '') + '/webhook'
  : (process.env.WEBHOOK_URL || '');

async function start() {
  if (WEBHOOK_URL) {
    const app = express();
    app.use(express.json({ limit: '1mb' }));
    app.get('/', (req, res) => res.send('Bot is alive'));
    app.post('/webhook', bot.webhookCallback('/webhook'));

    await bot.telegram.setWebhook(WEBHOOK_URL);
    app.listen(PORT, () => {
      console.log('Webhook server started on port ' + PORT);
      console.log('Webhook URL: ' + WEBHOOK_URL);
    });
    return;
  }
  await launchWithRetry();
}

start();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));


}

// Happ использует локальный SOCKS-прокси 127.0.0.1:10808 (см. config.json).
// Если у тебя другой прокси — задай его переменной BOT_PROXY в run_bot.bat.
// На Windows используем локальный прокси Happ. На Linux-хостинге прокси не нужен.
const HAPP_PROXY = process.platform === 'win32' ? 'socks://127.0.0.1:10808' : '';
const PROXY = process.env.BOT_PROXY || HAPP_PROXY || detectProxy();
let agent;
if (PROXY) {
  agent = /^socks/i.test(PROXY) ? new SocksProxyAgent(PROXY) : new HttpsProxyAgent(PROXY);
  console.log('Использую прокси: ' + PROXY);
} else {
  console.log('Прокси не найден — Trying direct connection.');
}
const netOpts = agent ? { agent, httpsAgent: agent } : {};

const bot = new Telegraf(TOKEN, agent ? { telegram: { agent } } : {});
const WIKI = 'https://ru.wikipedia.org/w/api.php';
const WD_API = 'https://www.wikidata.org/w/api.php';
const DDG = 'https://api.duckduckgo.com/';
const SPARQL = 'https://query.wikidata.org/sparql';

// Википедия отклоняет запросы без нормального User-Agent — из-за этого был 403.
const UA = 'SchoolBot/1.0 (school project; t.me/schol_spg_bot)';
axios.defaults.headers.common['User-Agent'] = UA;
axios.defaults.headers.common['Accept'] = 'application/json';

const TEXT_BTN = '📝 Готовый текст';
const PRES_BTN = '📊 Презентация';

const mainMenu = Markup.keyboard([[TEXT_BTN, PRES_BTN]]).resize();
const removeKb = Markup.removeKeyboard();

const YES_WORDS = ['да', 'давай', 'давай же', 'ага', 'угу', 'конечно', 'ок', 'окей',
  'ладно', 'хочу', 'сделай', 'презентацию', 'презентация', 'yes', 'yeah', 'добре'];
const NO_WORDS = ['нет', 'не надо', 'не хочу', 'не сейчас', 'позже', 'стоп', 'отмена', 'no'];

const yesNoKeyboard = () => Markup.inlineKeyboard([
  [Markup.button.callback('Да, сделать презентацию', 'pres:yes'),
   Markup.button.callback('Нет, спасибо', 'pres:no')],
]);

const STYLES = { classic: 'Классический', modern: 'Современный', kids: 'Детский' };
const COLORS = {
  blue: ['Синий', '1F4E8C'],
  green: ['Зелёный', '226E4E'],
  red: ['Красный', '962D37'],
  violet: ['Фиолетовый', '5E378C'],
};

const styleKeyboard = () => Markup.inlineKeyboard(
  Object.entries(STYLES).map(([k, label]) => Markup.button.callback(label, `style:${k}`))
);
const colorKeyboard = () => Markup.inlineKeyboard(
  Object.entries(COLORS).map(([k, v]) => Markup.button.callback(v[0], `color:${k}`))
);

// Простейшие сессии в памяти
const sessions = new Map();
function session(ctx) {
  const id = ctx.chat.id;
  if (!sessions.has(id)) sessions.set(id, { step: 'menu' });
  return sessions.get(id);
}

// ---------- Открытые данные ----------

// --- Понимание разговорных запросов ("напиши про энштейна" и т.п.) ---
const STOP_WORDS = new Set([
  'напиши','написать','напишите','про','об','about','сделай','сделать','сделайте',
  'давай','дай','пришли','инфа','инфу','информация','информацию',
  'текст','доклад','доклада','презентация','презентацию','презентации',
  'биография','биографию','биографии','расскажи','рассказать','кто','такой','такая',
  'что','какой','какая','нужно','надо','помоги','найди','ищи',
  'поищи','сведения','данные','материал','материалы','тему','тема','для','класс',
  'класса','школе','школа','школь','человек','личность','жизнь','жизни'
]);

// Частые сокращения и опечатки -> полное имя.
const ALIASES = {
  'энштейн': 'альберт эйнштейн', 'энштейна': 'альберт эйнштейн', 'айнштейн': 'альберт эйнштейн',
  'пушкин': 'александр пушкин', 'пушкина': 'александр пушкин',
  'лермонтов': 'михаил лермонтов', 'толстой': 'лев толстой', 'чехов': 'антон чехов',
  'достоевский': 'фёдор достоевский', 'гоголь': 'николай гоголь',
  'черчилль': 'уинстон черчилль', 'наполеон': 'наполеон бонапарт',
  'сталин': 'иосиф сталин', 'ленин': 'владимир ленин', 'гитлер': 'адольф гитлер',
  'петр 1': 'пётр великий', 'петр великий': 'пётр великий', 'пётр 1': 'пётр великий',
  'екатерина 2': 'екатерина великая', 'екатерина великая': 'екатерина великая',
  'гагарин': 'юрий гагарин', 'королев': 'сергей королёв', 'менделеев': 'дмитрий менделеев',
  'ломоносов': 'михаил ломоносов', 'суворов': 'александр суворов', 'жуков': 'георгий жуков'
};

function cleanQuery(raw) {
  let s = String(raw || '').toLowerCase().replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9\s-]/gi, ' ').replace(/\s+/g, ' ').trim();
  let words = s.split(' ').filter(w => w.length > 1 && !STOP_WORDS.has(w));
  s = words.join(' ');
  return ALIASES[s] || s || String(raw || '').trim();
}

function lev(a, b) {
  a = String(a || ''); b = String(b || '');
  const prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let last = prev[0]; prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const old = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, last + (a[i - 1] === b[j - 1] ? 0 : 1));
      last = old;
    }
  }
  return prev[b.length];
}

function sim(a, b) {
  a = String(a || '').toLowerCase(); b = String(b || '').toLowerCase();
  if (!a || !b) return 0;
  return 1 - lev(a, b) / Math.max(a.length, b.length);
}

function scoreCandidate(q, title, desc) {
  const t = String(title || '').toLowerCase().replace(/ё/g, 'е');
  const words = q.split(' ').filter(Boolean);
  let best = 0;
  for (const w of words) {
    if (t.includes(w)) best = Math.max(best, 0.92);
    for (const tw of t.split(' ')) best = Math.max(best, sim(w, tw) * 0.9);
    best = Math.max(best, sim(w, t) * 0.85);
  }
  best = Math.max(best, sim(q, t));
  if (desc && /человек|личность|персон/i.test(desc)) best += 0.04;
  return best;
}

async function wikiSearchRaw(name, limit = 8) {
  const url = WIKI + '?' + new URLSearchParams({
    action: 'query', list: 'search', srsearch: name,
    format: 'json', origin: '*', srlimit: String(limit),
  });
  const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
  return data.query?.search || [];
}

async function wdSearch(name) {
  const url = WD_API + '?' + new URLSearchParams({
    action: 'wbsearchentities', search: name, language: 'ru',
    format: 'json', origin: '*', limit: '10',
  });
  const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
  return data.search || [];
}

async function ddgEntity(name) {
  const { data } = await axios.get(DDG, {
    params: { q: name, format: 'json', no_html: 1, no_redirect: 1 },
    timeout: 20000, ...netOpts,
  });
  return data.Heading ? { title: data.Heading, desc: data.Abstract || '' } : null;
}

async function findPerson(raw) {
  let q = cleanQuery(raw);
  const queries = [q];
  if (q !== String(raw).trim()) queries.push(String(raw).trim());

  let candidates = [];
  for (const query of queries) {
    const [wiki, wd] = await Promise.allSettled([wikiSearchRaw(query, 8), wdSearch(query)]);
    if (wiki.status === 'fulfilled') {
      for (const h of wiki.value) {
        candidates.push({
          title: h.title, pageid: h.pageid,
          desc: String(h.snippet || '').replace(/<[^>]+>/g, ''),
          score: scoreCandidate(q, h.title, h.snippet || ''), source: 'wikipedia',
        });
      }
    }
    if (wd.status === 'fulfilled') {
      for (const e of wd.value) {
        candidates.push({
          title: e.label, qid: e.id, desc: e.description || '',
          score: scoreCandidate(q, e.label + ' ' + (e.aliases || []).join(' '), e.description || ''),
          source: 'wikidata',
        });
      }
    }
    if (candidates.length) break;

    // Если Википедия/Викиданные ничего не нашли — пробуем «яндексоподобный» ответ.
    try {
      const dd = await ddgEntity(query);
      if (dd) {
        const more = await wikiSearchRaw(dd.title, 3);
        for (const h of more) {
          candidates.push({
            title: h.title, pageid: h.pageid,
            desc: dd.desc || String(h.snippet || '').replace(/<[^>]+>/g, ''),
            score: scoreCandidate(q, h.title, dd.desc || ''), source: 'duckduckgo',
          });
        }
      }
    } catch (e) {}
    if (candidates.length) break;
  }

  const seen = new Set();
  candidates = candidates.filter(c => {
    const k = String(c.title || '').toLowerCase().replace(/ё/g, 'е');
    if (!k || seen.has(k)) return false;
    seen.add(k); return true;
  }).sort((a, b) => b.score - a.score);

  return { query: q, best: candidates[0] || null, all: candidates.slice(0, 5) };
}

async function wikiArticle(pageid) {
  const url = WIKI + '?' + new URLSearchParams({
    action: 'query', prop: 'extracts|pageprops', pageids: String(pageid),
    explaintext: '1', format: 'json', origin: '*',
  });
  const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
  const page = data.query?.pages?.[String(pageid)] || {};
  return {
    title: page.title || '',
    text: page.extract || '',
    qid: page.pageprops?.wikibase_item || '',
  };
}

async function sparql(q) {
  const { data } = await axios.get(SPARQL, {
    params: { query: q, format: 'json' }, timeout: 25000, ...netOpts,
  });
  return data.results?.bindings || [];
}

const val = (row, key) => row?.[key]?.value || '';

async function collect(raw) {
  const found = await findPerson(raw);
  if (!found.best) return null;

  let best = found.best;
  let pageid = best.pageid;

  if (!pageid && best.qid) {
    try {
      const url = WD_API + '?' + new URLSearchParams({
        action: 'wbgetentities', ids: best.qid, props: 'sitelinks',
        sitefilter: 'ruwiki', format: 'json', origin: '*',
      });
      const { data } = await axios.get(url, { timeout: 25000, ...netOpts });
      const ruTitle = data.entities?.[best.qid]?.sitelinks?.ruwiki?.title;
      if (ruTitle) {
        const wikiHits = await wikiSearchRaw(ruTitle, 3);
        pageid = wikiHits[0]?.pageid;
      }
    } catch (e) {}
    if (!pageid) {
      try {
        const wikiHits = await wikiSearchRaw(best.title, 3);
        pageid = wikiHits[0]?.pageid;
      } catch (e) {}
    }
  }
  if (!pageid) return null;

  const article = await wikiArticle(pageid);
  const qid = article.qid || best.qid || '';
  let facts = {}, family = {};

  if (qid) {
    const fRows = await sparql(`
      SELECT ?dob ?dod ?bornLabel ?diedLabel ?countryLabel ?jobLabel WHERE {
        VALUES ?person { wd:${qid} }
        OPTIONAL { ?person wdt:P569 ?dob }
        OPTIONAL { ?person wdt:P570 ?dod }
        OPTIONAL { ?person wdt:P19 ?born }
        OPTIONAL { ?person wdt:P20 ?died }
        OPTIONAL { ?person wdt:P27 ?country }
        OPTIONAL { ?person wdt:P106 ?job }
        SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
      } LIMIT 1`);
    const r = fRows[0] || {};
    facts = {
      birth: (val(r, 'dob') || '—').slice(0, 4),
      death: (val(r, 'dod') || '—').slice(0, 4),
      born: val(r, 'bornLabel') || '—',
      died: val(r, 'diedLabel') || '—',
      country: val(r, 'countryLabel') || '—',
      job: val(r, 'jobLabel') || '—',
    };

    const famRows = await sparql(`
      SELECT ?kind ?relLabel WHERE {
        VALUES ?person { wd:${qid} }
        VALUES (?p ?kind) {
          (wdt:P22 "Отец") (wdt:P25 "Мать") (wdt:P26 "Супруг(а)") (wdt:P40 "Дети")
        }
        ?person ?p ?rel .
        SERVICE wikibase:label { bd:serviceParam wikibase:language "ru,en". }
      } ORDER BY ?kind`);
    for (const row of famRows) {
      const k = val(row, 'kind');
      (family[k] ||= []).push(val(row, 'relLabel'));
    }
  }

  return {
    ...article, facts, family,
    query: found.query, matched: article.title, candidates: found.all,
    wikiUrl: 'https://ru.wikipedia.org/wiki/' + article.title.replace(/ /g, '_'),
    wdUrl: qid ? `https://www.wikidata.org/wiki/${qid}` : '',
  };
}

function sentences(text, limit = 6) {
  return String(text).replace(/\n/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .map(s => s.trim())
    .filter(s => s.length > 40)
    .slice(0, limit);
}

// ---------- Текст доклада ----------

function makeReport(d, note) {
  const f = d.facts;
  const lines = [];
  lines.push(`Доклад на тему: ${d.title}`, '');
  lines.push('Введение');
  lines.push(`${d.title} — известная личность. Родился(ась) в ${f.birth} году в городе ${f.born}. Гражданство: ${f.country}. Род деятельности: ${f.job}.`);
  lines.push('', 'Основная часть');
  for (const s of sentences(d.text, 6)) lines.push(s);
  lines.push('', 'Семья');
  const fam = Object.entries(d.family);
  if (fam.length) for (const [k, v] of fam) lines.push(`${k}: ${v.join(', ')}`);
  else lines.push('Сведения о семье в открытых источниках отсутствуют.');
  lines.push('', 'Заключение');
  lines.push(`${d.title} оставил(а) заметный след. Его/её биография изучается в школе и может быть использована для самостоятельной работы.`);
  if (note) lines.push('', 'Дополнительно учесть: ' + note);
  lines.push('', 'Источники:', d.wikiUrl);
  if (d.wdUrl) lines.push(d.wdUrl);
  return lines.join('\n');
}

// ---------- Презентация ----------

async function makePptx(d, styleKey, colorKey, filePath) {
  const color = COLORS[colorKey][1];
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  function addSlide(title, bullets, fontSize = 20) {
    const slide = pptx.addSlide();
    slide.background = { color };
    slide.addText(title, {
      x: 0.7, y: 0.5, w: 12, h: 1, fontSize: styleKey === 'kids' ? 38 : 34,
      bold: true, color: 'FFFFFF', align: 'center',
    });
    const text = bullets.map(b => '• ' + b).join('\n');
    slide.addText(text, {
      x: 0.9, y: 1.8, w: 11.6, h: 5.2,
      fontSize: styleKey === 'kids' ? fontSize + 2 : fontSize,
      color: 'FFFFFF', breakLine: false, paraSpaceAfter: 10,
    });
  }

  const f = d.facts;
  addSlide(d.title, ['Школьная презентация', `Стиль: ${STYLES[styleKey]}`, `Цвет: ${COLORS[colorKey][0]}`], 24);
  addSlide('Кто это?', [
    `Родился(ась): ${f.birth}`,
    `Место рождения: ${f.born}`,
    `Гражданство: ${f.country}`,
    `Деятельность: ${f.job}`,
  ]);
  addSlide('Биография', sentences(d.text, 4), 17);
  const fam = Object.entries(d.family).map(([k, v]) => `${k}: ${v.join(', ')}`);
  addSlide('Семья', fam.length ? fam : ['Сведений нет']);
  addSlide('География', [
    `Рождение: ${f.born}`, `Смерть: ${f.died}`, `Страна: ${f.country}`,
  ]);
  addSlide('Интересные факты', sentences(d.text, 3).reverse(), 17);
  addSlide('Вывод', [
    `${d.title} — важная историческая личность.`,
    'Его/её жизнь связана с ключевыми событиями эпохи.',
    'Изучение биографии помогает лучше понять историю.',
  ]);
  addSlide('Источники', [d.wikiUrl, d.wdUrl || 'Викиданные'].filter(Boolean), 16);

  await pptx.writeFile({ fileName: filePath });
}

// ---------- Хендлеры ----------

function matchesWord(text, words) {
  const low = String(text || '').toLowerCase().replace(/ё/g, 'е').trim();
  return words.some(w => low === w || low.startsWith(w + ' ') || low.includes(' ' + w + ' '));
}

async function askStyle(ctx, s) {
  s.step = 'pres_style';
  await ctx.reply('Выбери стиль презентации:', styleKeyboard());
}

async function startPresentationForLast(ctx, s) {
  if (s.lastName) {
    s.name = s.lastName;
    s.note = s.lastNote || '';
    await ctx.reply(`Делаю презентацию про: ${s.lastName}`);
    await askStyle(ctx, s);
  } else {
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
  }
}

async function handleTextRequest(ctx, s, text) {
  const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
  if (!parts.length) return ctx.reply('Напиши, про кого сделать работу.');
  await ctx.reply('Понял запрос. Ищу и сверяю по нескольким источникам...');
  try {
    const d = await collect(parts[0]);
    if (!d) {
      return ctx.reply(
        'Не смог понять, о ком речь. Напиши имя и фамилию полностью, ' +
        'например: Альберт Эйнштейн. Или просто: напиши про Эйнштейна.'
      );
    }
    if (d.matched && d.query !== d.matched.toLowerCase()) {
      await ctx.reply(`Похоже, ты имел(а) в виду: ${d.matched}`);
    }
    const note = parts.slice(1).join(' ');
    const report = makeReport(d, note);
    for (let i = 0; i < report.length; i += 3500) {
      await ctx.reply(report.slice(i, i + 3500));
    }
    s.lastName = d.matched || parts[0];
    s.lastNote = note;
    s.awaitingPres = true;
    s.step = 'menu';
    await ctx.reply('Хочешь также презентацию?', yesNoKeyboard());
  } catch (e) {
    console.error(e);
    await ctx.reply('Ошибка: ' + e.message);
  }
}

bot.start(async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply(
    'Привет! Я помогу с школьной работой про человека.\n\nВыбери, что нужно сделать:',
    mainMenu
  );
});

bot.command('help', async (ctx) => {
  await ctx.reply(
    '📝 Готовый текст — напиши, про кого, получишь текст для вставки.\n' +
    '📊 Презентация — напиши тему, ответь на вопросы про дизайн и получи файл.\n\n' +
    '/start — начать заново'
  );
});

bot.command('cancel', async (ctx) => {
  sessions.set(ctx.chat.id, { step: 'menu' });
  await ctx.reply('Отменено. /start — начать заново.', mainMenu);
});

bot.on('text', async (ctx) => {
  const text = ctx.message.text.trim();
  const s = session(ctx);

  if (s.awaitingPres && s.step === 'menu') {
    if (matchesWord(text, YES_WORDS)) {
      s.awaitingPres = false;
      await startPresentationForLast(ctx, s);
      return;
    }
    if (matchesWord(text, NO_WORDS)) {
      s.awaitingPres = false;
      await ctx.reply('Хорошо. Если передумаешь — нажми 📊 Презентация.', mainMenu);
      return;
    }
    s.awaitingPres = false;
  }

  if (text === TEXT_BTN) {
    s.step = 'text_topic';
    await ctx.reply(
      'Напиши, про кого сделать текст, и что учесть.\nНапример:\nУинстон Черчилль\nДля 7 класса, про семью и войну',
      removeKb
    );
    return;
  }

  if (text === PRES_BTN) {
    s.awaitingPres = false;
    s.step = 'pres_topic';
    await ctx.reply(
      'Напиши, про кого презентация и что должно быть на слайдах.\nНапример:\nАльберт Эйнштейн\nБиография, семья, открытия, вывод',
      removeKb
    );
    return;
  }

  if (s.step === 'text_topic') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  if (s.step === 'pres_topic') {
    const parts = text.split('\n').map(x => x.trim()).filter(Boolean);
    if (!parts.length) return ctx.reply('Напиши имя человека.');
    s.name = parts[0];
    s.note = parts.slice(1).join(' ');
    s.step = 'pres_style';
    await ctx.reply('Выбери стиль презентации:', styleKeyboard());
    return;
  }

  // Если человек просто написал «напиши про энштейна» — сразу понимаем это как запрос текста.
  if (s.step === 'menu') {
    await handleTextRequest(ctx, s, text);
    return;
  }

  await ctx.reply('Выбери кнопку или напиши /start.', mainMenu);
});

bot.action('pres:yes', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Отлично, делаю презентацию.');
  s.awaitingPres = false;
  await startPresentationForLast(ctx, s);
});

bot.action('pres:no', async (ctx) => {
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText('Хорошо. Если передумаешь — нажми 📊 Презентация.');
  s.awaitingPres = false;
  s.step = 'menu';
  await ctx.reply('Что ещё сделать?', mainMenu);
});

bot.action(/^style:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  s.style = key;
  s.step = 'pres_color';
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Стиль: ${STYLES[key]}. Теперь выбери цвет:`);
  await ctx.reply('Цвет:', colorKeyboard());
});

bot.action(/^color:(.+)$/, async (ctx) => {
  const key = ctx.match[1];
  const s = session(ctx);
  await ctx.answerCbQuery();
  await ctx.editMessageText(`Цвет: ${COLORS[key][0]}. Делаю презентацию...`);
  try {
    const d = await collect(s.name);
    if (!d) return ctx.reply('Не нашёл человека в Википедии.');
    const filePath = path.join(os.tmpdir(), `presentation_${ctx.chat.id}.pptx`);
    await makePptx(d, s.style, key, filePath);
    await ctx.replyWithDocument(
      { source: fs.createReadStream(filePath), filename: `Презентация_${d.title}.pptx` },
      { caption: `Готовая презентация: ${d.title}\nСтиль: ${STYLES[s.style]}, цвет: ${COLORS[key][0]}` }
    );
    fs.unlinkSync(filePath);
    await ctx.reply('Ещё что-нибудь?', mainMenu);
    s.step = 'menu';
  } catch (e) {
    console.error(e);
    await ctx.reply('Не удалось сделать презентацию: ' + e.message);
    await ctx.reply('/start — вернуться в меню', mainMenu);
    s.step = 'menu';
  }
});

async function launchWithRetry() {
  for (let i = 1; i <= 12; i++) {
    try {
      await bot.launch();
      console.log('Бот запущен (polling)');
      return;
    } catch (e) {
      console.log(`Попытка ${i}/12 не удалась: ${e.message}. Повтор через 5 сек...`);
      await new Promise(r => setTimeout(r, 5000));
    }
  }
  console.error('Не удалось подключиться к Telegram. Проверь, что Happ подключён и прокси 10808 работает.');
  process.exit(1);
}

// --- Облако: webhook вместо polling ---
const PORT = process.env.PORT || 3000;
const WEBHOOK_URL = process.env.RENDER_EXTERNAL_URL
  ? process.env.RENDER_EXTERNAL_URL.replace(/\/$/, '') + '/webhook'
  : (process.env.WEBHOOK_URL || '');

async function start() {
  if (WEBHOOK_URL) {
    const app = express();
    app.use(express.json({ limit: '1mb' }));
    app.get('/', (req, res) => res.send('Bot is alive'));
    app.post('/webhook', bot.webhookCallback('/webhook'));

    await bot.telegram.setWebhook(WEBHOOK_URL);
    app.listen(PORT, () => {
      console.log('Webhook server started on port ' + PORT);
      console.log('Webhook URL: ' + WEBHOOK_URL);
    });
    return;
  }
  await launchWithRetry();
}

start();

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));

