import * as React from 'react';
import { ICrudProps } from './ICrudProps';
import { getSP, SPFI } from '../../../pnpjsConfig';
import { PrimaryButton, DetailsList, SelectionMode, IconButton, Dialog, DialogType, DialogFooter, TextField, DefaultButton } from 'office-ui-fabric-react';

// type definitions
interface IQuoteRes {
  Title: string;
  Author0: string;
  Id: number;
}

interface IQuote {
  quote: string;
  author: string;
  id: number;
}

const Crud = (props: ICrudProps): React.ReactElement => {
  const _sp: SPFI = getSP(props.spcontext);
  const [reload, setReload] = React.useState<boolean>(false);
  const [quotes, setQuotes] = React.useState<Array<IQuote>>([]);
  const [currentId, setCurrentId] = React.useState<number | any>();
  const [isEditHidden, setIsEditHidden] = React.useState<boolean>(true);
  const [editedQuote, setEditedQuote] = React.useState<string>('');
  const [editedAuthor, setEditedAuthor] = React.useState<string>('');
  const [isAddHidden, setIsAddHidden] = React.useState<boolean>(true);
  const [newQuote, setNewQuote] = React.useState<string>('');
  const [newAuthor, setNewAuthor] = React.useState<string>('');

  // use effect hook to call function each time page reloads
  React.useEffect(() => {
    getListItems();
  }, [reload])

  const getListItems = async () => {
    //this function gets list items from the site specified in the serve.json file 
    try {
      //fetching the list items
      const getListItems = await _sp.web.lists.getByTitle('Quotes').items();
      //setting the list items to a state variable
      setQuotes(getListItems.map((each: IQuoteRes) => ({
        quote: each.Title,
        author: each.Author0,
        id: each.Id
      })))
    } catch (e) {
      //log any errors if there are any
      console.log(e);
    } finally {
      console.log('List items fetched', quotes);
    }
  }

  const handleQuote = (event: React.ChangeEvent<HTMLInputElement>) => {
    setNewQuote(event.target.value);
  };

  const handleAuthor = (event: React.ChangeEvent<HTMLInputElement>) => {
    setNewAuthor(event.target.value);
  };

  const addNewListItem = async () => {
    // Get a reference to the SharePoint list named "Quotes"
    const list = _sp.web.lists.getByTitle("Quotes");
    try {
      // Add a new item to the list with the provided values
      await list.items.add({
        Title: newQuote,
        Author0: newAuthor
      });
      // Close the add modal/dialog
      setIsAddHidden(true);
      // Trigger a reload by toggling the 'reload' state variable
      setReload(!reload);
      // Log a message to indicate that the list item has been successfully added
      console.log('List item added');
    } catch (e) {
      // Log any errors that occur during the addition process
      console.log(e);
    } finally {
      // Ensure that the add modal/dialog is closed, even in case of an error
      setIsAddHidden(true);
    }
  }

  const openEditDialog = (id: number) => {
    setCurrentId(id)
    //this function would open the edit dialogue and expose a form
    setIsEditHidden(false);
    const quote: IQuote | undefined = quotes.find((each) => each.id === id); // you may have to make changes to your tsconfig.json to use find, "target": "es2015"
    // setting form with previous quote so user can edit
    if (quote) {
      setEditedAuthor(quote.author);
      setEditedQuote(quote.quote)
    }
  };

  const handleQuoteChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    //handling change of quote
    setEditedQuote(event.target.value);
  };

  const handleAuthorChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    //handling change of author
    setEditedAuthor(event.target.value);
  };

  const editListItem = async () => {
    // Get a reference to the SharePoint list named "Quotes"
    const list = _sp.web.lists.getByTitle("Quotes");
    try {
      // Update the list item with the currentId using the provided values
      await list.items.getById(currentId).update({
        Title: editedQuote,
        Author0: editedAuthor
      });
      // Close the edit modal/dialog
      setIsEditHidden(true);
      // Trigger a reload by toggling the 'reload' state variable
      setReload(!reload);
      // Log a message to indicate that the list item has been successfully edited
      console.log('List item edited');
    } catch (e) {
      // Log any errors that occur during the update process
      console.log(e);
    } finally {
      // Ensure that the edit modal/dialog is closed, even in case of an error
      setIsEditHidden(true);
    }
  }


  const deleteListItem = async (id: number) => {
    // Get a reference to the SharePoint list named "Quotes"
    const list = _sp.web.lists.getByTitle("Quotes");
    try {
      // Delete the list item with the specified ID
      await list.items.getById(id).delete();
      // Trigger a reload by toggling the 'reload' state variable
      setReload(!reload);
      // Log a message to indicate that the list item has been successfully deleted
      console.log('List item deleted');
    } catch (e) {
      // Log any errors that occur during the deletion process
      console.log(e);
    }
  }
  

  return (
    <>
      <h1>Performing CRUD Operations on SharePoint List Items with PnPjs</h1>
      <div className='quoteBox'>
        <h2>Quotes</h2>
        <div className='quoteContainer'>
          <DetailsList
            items={quotes || []}
            columns={[
              {
                key: 'quoteColumn',
                name: 'Quote',
                fieldName: 'quote',
                minWidth: 200,
                isResizable: true,
                onRender: (item: IQuote) => <div>{item.quote}</div>,
              },
              {
                key: 'authorColumn',
                name: 'Author',
                fieldName: 'author',
                minWidth: 100,
                isResizable: true,
                onRender: (item: IQuote) => <div>{item.author}</div>,
              },
              {
                key: 'actionsColumn',
                name: 'Actions',
                minWidth: 100,
                isResizable: true,
                onRender: (item: IQuote) => (
                  <div>
                    <IconButton
                      iconProps={{ iconName: 'Edit' }}
                      // onClick={() => editListItem(item.id)} // handles the edit list item functionality
                      onClick={() => openEditDialog(item.id)}
                      title="Edit"
                      ariaLabel="Edit"
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      onClick={() => deleteListItem(item.id)} // handles the delete list item functionality
                      title="Delete"
                      ariaLabel="Delete"
                    />
                  </div>
                ),
              },
            ]}
            selectionMode={SelectionMode.none}
          />
          <Dialog
            hidden={isEditHidden}
            onDismiss={() => setIsEditHidden(true)}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Edit Quote',
            }}
          >
            <div>
              <TextField
                label="Quote"
                value={editedQuote}
                onChange={handleQuoteChange}
              />
              <TextField
                label="Author"
                value={editedAuthor}
                onChange={handleAuthorChange}
              />
            </div>
            <DialogFooter>
              <PrimaryButton text="Submit" onClick={() => editListItem()} />
              <DefaultButton text="Cancel" onClick={() => setIsEditHidden(true)} />
            </DialogFooter>
          </Dialog>
        </div>
        <div>
          <PrimaryButton text='Add Quote' />
        </div>
        <Dialog
          hidden={isAddHidden}
          onDismiss={() => setIsAddHidden(true)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Add Quote',
          }}
        >
          <div>
            <TextField
              label="Quote"
              value={newQuote}
              onChange={handleQuote}
            />
            <TextField
              label="Author"
              value={newAuthor}
              onChange={handleAuthor}
            />
          </div>
          <DialogFooter>
            <PrimaryButton text="Submit" onClick={() => addNewListItem()} />
            <DefaultButton text="Cancel" onClick={() => setIsAddHidden(true)} />
          </DialogFooter>
        </Dialog>
      </div>
    </>
  )
}
export default Crud;