const { Client, APIErrorCode } = require("@notionhq/client")

let notion = new Client({ auth: process.env.NOTION_TOKEN });
let users = new Array();

function prettyPage(page)
{
    let json = {}
    const prop = page.properties;
    for(let key in prop) {
      let d = page.properties[key][page.properties[key].type];
      console.log(page);
      if(!d) continue;
      
      let pretty = function(data) {
        if(page.properties[key].type == "select" || page.properties[key].type == "multi_select"){
          data.type = "select"
        }
        switch(data.type){
        case "text":
          data = data.plain_text
          break;
        case "select":
          data = data.name
          break;
        case "person":
          data = {
            name: data.name,
            avatar_url: data.avatar_url,
            data: data.person
          }
          break;
        case "email":
        case "Email":
        default:
          break;
        }
        return data;
      }
      
      let dt = null;
      if(!d.length){
        dt = pretty(d)
      }else if(typeof d == "string"){
        dt = pretty(d)
      }else {
        if(d.length > 1 || page.properties[key].type == "select" || page.properties[key].type == "multi_select") {
          dt = [];
          for(let i=0; i<d.length; ++i) {
            dt[i] = pretty(d[i])
          }
        }else if(d.length == 1){
          dt = pretty(d[0])
        }
      }
      
      json[key] = dt;
    };
    return json;
}

async function getNotionUsers() {
  let cursor = undefined;
  const ret = []
  
  while (true) {
    const { results, next_cursor } = await notion.users.list({
      start_cursor: cursor
    });
        
    results.forEach(u => {
      if(!users.includes(u.id)) {
        users.push(u);
      }else{
        console.log("find:" + u.id);
      }
    });
    
    ret.push(...results);
    
    if (!next_cursor) {
      break
    }
    cursor = next_cursor
  }
  
  return ret;
}

async function getSettingFromNotion(dbId, filter) {
  const config = [];
  let cursor = undefined;
  
  while (true) {
    let nc = undefined;
    
    if(filter){
      const { results, next_cursor } = await notion.databases.query({
        database_id: dbId,
        filter: filter,
        start_cursor: cursor
      });
      config.push(...results);
      nc = next_cursor;
    }else{
      const { results, next_cursor } = await notion.databases.query({
        database_id: dbId,
        start_cursor: cursor
      });
      config.push(...results);
      nc = next_cursor;
    }
    
    if (!nc) {
      break;
    }
    
    cursor = nc
  }
  
  return config.map(page => {
    let json = prettyPage(page);
    
    //必要な情報を設定しとく、かぶったら知らん…
    json.child_id = page.id;
    json.page_type = page.object;
    json.url = page.url;
      
    return json;
  });
}

async function getChildBlocks(pgId) {
  const contents = [];
  let cursor = undefined;
  while (true) {
    const { results, next_cursor } = await notion.blocks.children.list({
    	block_id: pgId
    });
    contents.push(...results);
    if (!next_cursor) {
      break
    }
    cursor = next_cursor
  }
  return contents;
};

async function getPageContents(pgId, withChild) {
  let recursive_get_childs = async function(cts)
  {
    for(let i=0; i<cts.length; ++i)
    {
      let block = cts[i];
      if(block.has_children)
      {
        let children = await getChildBlocks(block.id);
        await recursive_get_childs(children);
        block.children = children;
      }
    }
  }
  
  let contents = await getChildBlocks(pgId);
  if(withChild)
  {
    await recursive_get_childs(contents);
  }
  return contents;
}

async function getNotionDatabase(dbId, filter) {
  const pages = []
  let cursor = undefined;
  
  while (true) {
    let nc = undefined;
    if(filter){
      const { results, next_cursor } = await notion.databases.query({
        database_id: dbId,
        filter: filter,
        start_cursor: cursor
      });
      pages.push(...results);
      nc = next_cursor;
    }else{
      const { results, next_cursor } = await notion.databases.query({
        database_id: dbId,
        start_cursor: cursor
      });
      pages.push(...results);
      nc = next_cursor;
    }
    
    if (!nc) {
      break
    }
    
    cursor = nc
  }
  
  return pages;
}

async function searchNotionDatabases(query, filter) {
  const pages = [];
  let cursor = undefined;
  
  while (true) {
    let nc = undefined;
    if(filter){
      const { results, next_cursor } = await notion.search({
        query: query,
        filter: filter,
        start_cursor: cursor,
        page_size: 50
      });
      pages.push(...results);
      nc = next_cursor;
    }else{
      const { results, next_cursor } = await notion.search({
        query: query,
        start_cursor: cursor,
        page_size: 50
      });
      pages.push(...results);
      nc = next_cursor;
    }
    
    if (!nc) {
      break
    }
    
    cursor = nc
  }
  
  return pages;
}

exports.getUserList = async () =>
{
	if(users.length == 0){
		await getNotionUsers();
	}
	return users;
}

exports.getConfig = async (dbId, filter) =>
{
	let config = await getSettingFromNotion(dbId, filter);
	return config;
}

exports.getPageProperties = async (pgId) =>
{
	let propeties = await notion.pages.retrieve({
		page_id: pgId
	});
	return propeties;
}

exports.getPageConf = async (pgId) =>
{
	let page = await notion.pages.retrieve({
		page_id: pgId
	});
	return prettyPage(page);
}

exports.getContents = async (pgId, withChilds = false) => 
{
	let page = await getPageContents(pgId, withChilds);
	return page;
}

exports.getDatabase = async(dbId, filter) =>
{
	let pages = await getNotionDatabase(dbId, filter);
	return pages;
}

exports.searchDocs = async (query, filter) =>
{
	let pages = await searchNotionDatabases(query, filter);
	return pages;
}

exports.createPage = async (data) => 
{
  let response = await notion.pages.create(data);
  return response;
}


exports.convertToSlackBlock = (block, indent_level) => 
{
//  console.log(block)
  switch(block.type) {
  case "title":
    return { type: "header", "text": { type:"plain_text", "text": block.title[0].text.content, "emoji": true } };
  case "heading_1":
    return { type: "header", "text": { type:"plain_text", "text": block.heading_1.text[0].plain_text, "emoji": true } };
  case "heading_2":
  case "heading_3":
    return { type: "section", "text": { type:"plain_text", "text": block[block.type].text[0].plain_text, "emoji": true } };
  case "bulleted_list_item":
  case "paragraph":
    let text = "";
    for(let i=0; i<block[block.type].text.length; ++i){
      for(let j=0; j<indent_level; ++j){
        text += "　";
      }
      
      if(block.type == "bulleted_list_item")
      {
        text += "・";
      }
      
      if(block[block.type].text[i].href)
      {
        text += block[block.type].text[i].plain_text.replace(block[block.type].text[i].href, "<" + block[block.type].text[i].href + ">");
      }
      else
      {
        text += block[block.type].text[i].plain_text
      }
      text += "\n";
    }
    if(text == "") return null;
    return { type: "section", "text": { type:"mrkdwn", "text": text } };
  }
  /*
	"accessory": {
		"type": "button",
		"text": {
			"type": "plain_text",
			"text": "Click Me",
			"emoji": true
		},
		"value": "click_me_123",
		"url": "https://google.com",
		"action_id": "button-action"
			}
	*/
  return null;
}

exports.convertToSlackUser = (user) =>
{
  return {
    username: user.name,
    icon_url: user.avatar_url
  }
}

/*
(async function(){
	let config = await getNotionPages("2f4d5a0435b349e5ac21e6d38ba1da2a");
	console.log(config);
	return;
	
	let users = await getNotionUsers();
	console.log(users);
	return;
	
	let result = await getTasksFromNotionDatabase();
	
	console.log(JSON.stringify(result[0], null, 2));
	console.log(result[0]["作成者"].created_by);
	console.log(result[0]["Title"].title);
	console.log(result[0]["Title"].title);
	//let result2 = await getTasksFromNotionDatabase();
	//console.log(result);
})();
*/