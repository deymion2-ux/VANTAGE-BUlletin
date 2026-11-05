const fs=require('fs'),vm=require('vm'),assert=require('assert'),path=require('path');
class Node {
 constructor(name){ this.name=name;this.children=[];this.dataset={};this.classList={add(){},remove(){}};this.className='';this.textContent=''; }
 replaceChildren(...children){this.children=children}
 appendChild(ch){if(ch.name==='fragment') this.children.push(...ch.children); else this.children.push(ch);return ch}
 append(...children){this.children.push(...children)}
 addEventListener(){}
}
const nodes={};const ctx={
 console,Date,URL,AbortController,Intl,setInterval,clearInterval,setTimeout,clearTimeout,
 navigator:{userAgent:'Chrome'},localStorage:{getItem(){return null},setItem(){}},
 document:{addEventListener(){},createElement:n=>new Node(n),createDocumentFragment:()=>new Node('fragment'),getElementById:id=>nodes[id]??(nodes[id]=new Node(id))},
 window:{location:{origin:'https://vantage.test',pathname:'/index.html'},addEventListener(){},open(){}},
 fetch:async ()=>({ok:true,json:async()=>({value:[{subject:'<img src=x onerror=alert(1)>', start:{dateTime:'2026-11-05T11:00:00.0000000'},webLink:'https://outlook.office.com/calendar/item'}]})}),
};
ctx.window.msal={PublicClientApplication:class{getActiveAccount(){return {name:'user'}};getAllAccounts(){return [{name:'user'}]};setActiveAccount(){} }};
ctx.msal=ctx.window.msal;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname,'..','assets','app.js'),'utf8'),ctx);
vm.runInContext('currentUserAccount={name:"user"}; sessionEpoch=5;',ctx);
vm.runInContext('renderDirectoryUsers([{displayName:"<svg onload=alert(1)>",mail:"person@example.com"}])',ctx);
const dir=nodes['sidebar-presence-list'].children[0];
assert.equal(dir.children[0].children[0].textContent,'<svg onload=alert(1)>');
assert.equal(dir.children[0].children[1].textContent,'person@example.com');
assert.equal(dir.children[1].title,'Kein Live-Präsenzstatus verfügbar');
(async()=>{
 await vm.runInContext('fetchOutlookCalendars("fake",new AbortController().signal,5)',ctx);
 const calendarNodes=[...nodes['today-calendar-events'].children,...nodes['upcoming-calendar-events'].children];
 assert(calendarNodes.length>0);
 const line=calendarNodes.find(n=>n.children.length>0)?.children[0];
 assert(line);
 assert(line.textContent.includes('<img src=x onerror=alert(1)>'));
 const dangerous=vm.runInContext("safeOutlookUrl('https://outlook.office.com.evil.org/x')",ctx);
 assert.equal(dangerous,'https://outlook.office.com/calendar/');
 console.log('VM security smoke checks passed: Graph text nodes and Outlook host allowlist.');
})().catch(err=>{console.error(err);process.exitCode=1});
